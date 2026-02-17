use schemars::JsonSchema;
use serde::{Deserialize, Serialize};
use std::collections::BTreeMap;

const MAX_GROUPS: usize = 50;
const MAX_SAMPLE_ADDRESSES: usize = 5;
const FORMULA_PREVIEW_MAX_BYTES: usize = 80;

pub const FORMULA_PARSE_FAILED: &str = "FORMULA_PARSE_FAILED";
pub const FORMULA_PARSE_FAILED_PREFIX: &str = "formula parse failed: ";

#[derive(
    Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize, JsonSchema, clap::ValueEnum, Default,
)]
#[serde(rename_all = "snake_case")]
pub enum FormulaParsePolicy {
    /// Abort on any formula parse failure.
    Fail,
    /// Continue but collect diagnostics.
    #[default]
    Warn,
    /// Skip silently.
    Off,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize, JsonSchema)]
#[serde(rename_all = "snake_case")]
pub enum CommandClass {
    SingleWrite,
    BatchWrite,
    ReadAnalysis,
}

impl FormulaParsePolicy {
    pub fn default_for_command_class(class: CommandClass) -> Self {
        match class {
            CommandClass::SingleWrite => FormulaParsePolicy::Fail,
            CommandClass::BatchWrite | CommandClass::ReadAnalysis => FormulaParsePolicy::Warn,
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct FormulaParseErrorGroup {
    pub error_code: String,
    pub error_message: String,
    pub sheet_name: String,
    pub formula_preview: String,
    pub count: usize,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub sample_addresses: Vec<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct FormulaParseDiagnostics {
    pub policy: FormulaParsePolicy,
    pub total_errors: usize,
    pub groups_truncated: bool,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub groups: Vec<FormulaParseErrorGroup>,
}

pub struct FormulaParseDiagnosticsBuilder {
    policy: FormulaParsePolicy,
    groups: BTreeMap<(String, String, String), GroupAccumulator>,
    total_errors: usize,
}

struct GroupAccumulator {
    error_code: String,
    error_message: String,
    count: usize,
    sample_addresses: Vec<String>,
}

impl FormulaParseDiagnosticsBuilder {
    pub fn new(policy: FormulaParsePolicy) -> Self {
        Self {
            policy,
            groups: BTreeMap::new(),
            total_errors: 0,
        }
    }

    pub fn record_error(&mut self, sheet: &str, address: &str, formula: &str, error: &str) {
        let formula_preview = truncate_formula_preview(formula);
        let key = (sheet.to_string(), error.to_string(), formula_preview);

        self.total_errors += 1;

        let group = self.groups.entry(key).or_insert_with(|| GroupAccumulator {
            error_code: FORMULA_PARSE_FAILED.to_string(),
            error_message: error.to_string(),
            count: 0,
            sample_addresses: Vec::new(),
        });

        group.count += 1;
        if group.sample_addresses.len() < MAX_SAMPLE_ADDRESSES {
            group.sample_addresses.push(address.to_string());
        }
    }

    pub fn build(self) -> FormulaParseDiagnostics {
        let groups_truncated = self.groups.len() > MAX_GROUPS;
        let groups = self
            .groups
            .into_iter()
            .take(MAX_GROUPS)
            .map(
                |((sheet_name, _error_message, formula_preview), group)| FormulaParseErrorGroup {
                    error_code: group.error_code,
                    error_message: group.error_message,
                    sheet_name,
                    formula_preview,
                    count: group.count,
                    sample_addresses: group.sample_addresses,
                },
            )
            .collect();

        FormulaParseDiagnostics {
            policy: self.policy,
            total_errors: self.total_errors,
            groups_truncated,
            groups,
        }
    }

    pub fn is_empty(&self) -> bool {
        self.total_errors == 0
    }

    pub fn has_errors(&self) -> bool {
        self.total_errors > 0
    }
}

fn truncate_formula_preview(formula: &str) -> String {
    if formula.len() <= FORMULA_PREVIEW_MAX_BYTES {
        return formula.to_string();
    }

    let mut end = FORMULA_PREVIEW_MAX_BYTES;
    while end > 0 && !formula.is_char_boundary(end) {
        end -= 1;
    }

    let mut result = formula[..end].to_string();
    result.push('…');
    result
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_policy_default_is_warn() {
        assert_eq!(FormulaParsePolicy::default(), FormulaParsePolicy::Warn);
    }

    #[test]
    fn test_policy_default_for_single_write() {
        assert_eq!(
            FormulaParsePolicy::default_for_command_class(CommandClass::SingleWrite),
            FormulaParsePolicy::Fail
        );
    }

    #[test]
    fn test_policy_default_for_batch_write() {
        assert_eq!(
            FormulaParsePolicy::default_for_command_class(CommandClass::BatchWrite),
            FormulaParsePolicy::Warn
        );
    }

    #[test]
    fn test_policy_default_for_read_analysis() {
        assert_eq!(
            FormulaParsePolicy::default_for_command_class(CommandClass::ReadAnalysis),
            FormulaParsePolicy::Warn
        );
    }

    #[test]
    fn test_policy_serde_roundtrip() {
        let cases = [
            (FormulaParsePolicy::Fail, "fail"),
            (FormulaParsePolicy::Warn, "warn"),
            (FormulaParsePolicy::Off, "off"),
        ];

        for (policy, expected) in cases {
            let serialized = serde_json::to_string(&policy).expect("serialize policy");
            assert_eq!(serialized, format!("\"{expected}\""));

            let deserialized: FormulaParsePolicy =
                serde_json::from_str(&serialized).expect("deserialize policy");
            assert_eq!(deserialized, policy);
        }
    }

    #[test]
    fn test_empty_builder() {
        let builder = FormulaParseDiagnosticsBuilder::new(FormulaParsePolicy::Warn);
        assert!(builder.is_empty());
        assert!(!builder.has_errors());

        let diagnostics = builder.build();
        assert_eq!(diagnostics.policy, FormulaParsePolicy::Warn);
        assert_eq!(diagnostics.total_errors, 0);
        assert!(!diagnostics.groups_truncated);
        assert!(diagnostics.groups.is_empty());
    }

    #[test]
    fn test_single_error_group() {
        let mut builder = FormulaParseDiagnosticsBuilder::new(FormulaParsePolicy::Warn);
        builder.record_error("Sheet1", "A1", "=SUM(A:A)", "unexpected token");

        let diagnostics = builder.build();
        assert_eq!(diagnostics.total_errors, 1);
        assert_eq!(diagnostics.groups.len(), 1);
        let group = &diagnostics.groups[0];
        assert_eq!(group.count, 1);
        assert_eq!(group.error_code, FORMULA_PARSE_FAILED);
    }

    #[test]
    fn test_grouping_same_key() {
        let mut builder = FormulaParseDiagnosticsBuilder::new(FormulaParsePolicy::Warn);
        builder.record_error("Sheet1", "A1", "=SUM(A:A)", "unexpected token");
        builder.record_error("Sheet1", "A2", "=SUM(A:A)", "unexpected token");
        builder.record_error("Sheet1", "A3", "=SUM(A:A)", "unexpected token");

        let diagnostics = builder.build();
        assert_eq!(diagnostics.groups.len(), 1);
        let group = &diagnostics.groups[0];
        assert_eq!(group.count, 3);
        assert_eq!(group.sample_addresses, vec!["A1", "A2", "A3"]);
    }

    #[test]
    fn test_grouping_different_sheets() {
        let mut builder = FormulaParseDiagnosticsBuilder::new(FormulaParsePolicy::Warn);
        builder.record_error("A", "A1", "=SUM(A:A)", "unexpected token");
        builder.record_error("B", "A1", "=SUM(A:A)", "unexpected token");

        let diagnostics = builder.build();
        assert_eq!(diagnostics.groups.len(), 2);
        assert_eq!(diagnostics.groups[0].sheet_name, "A");
        assert_eq!(diagnostics.groups[1].sheet_name, "B");
    }

    #[test]
    fn test_grouping_different_messages() {
        let mut builder = FormulaParseDiagnosticsBuilder::new(FormulaParsePolicy::Warn);
        builder.record_error("Sheet1", "A1", "=SUM(A:A)", "unexpected token");
        builder.record_error("Sheet1", "B1", "=SUM(A:A)", "unknown function");

        let diagnostics = builder.build();
        assert_eq!(diagnostics.groups.len(), 2);
    }

    #[test]
    fn test_sample_address_cap_at_5() {
        let mut builder = FormulaParseDiagnosticsBuilder::new(FormulaParsePolicy::Warn);

        for i in 1..=8 {
            builder.record_error("Sheet1", &format!("A{i}"), "=SUM(A:A)", "unexpected token");
        }

        let diagnostics = builder.build();
        let group = &diagnostics.groups[0];
        assert_eq!(group.count, 8);
        assert_eq!(group.sample_addresses.len(), 5);
        assert_eq!(group.sample_addresses, vec!["A1", "A2", "A3", "A4", "A5"]);
    }

    #[test]
    fn test_deterministic_ordering() {
        let mut builder = FormulaParseDiagnosticsBuilder::new(FormulaParsePolicy::Warn);
        builder.record_error("C", "A1", "=1", "err");
        builder.record_error("A", "A1", "=1", "err");
        builder.record_error("B", "A1", "=1", "err");

        let diagnostics = builder.build();
        let sheets: Vec<&str> = diagnostics
            .groups
            .iter()
            .map(|group| group.sheet_name.as_str())
            .collect();
        assert_eq!(sheets, vec!["A", "B", "C"]);
    }

    #[test]
    fn test_groups_truncated_at_50() {
        let mut builder = FormulaParseDiagnosticsBuilder::new(FormulaParsePolicy::Warn);

        for i in 0..60 {
            builder.record_error(
                "Sheet1",
                "A1",
                &format!("=SOME_LONG_FORMULA_{i}"),
                "unexpected token",
            );
        }

        let diagnostics = builder.build();
        assert_eq!(diagnostics.total_errors, 60);
        assert_eq!(diagnostics.groups.len(), 50);
        assert!(diagnostics.groups_truncated);
    }

    #[test]
    fn test_formula_preview_truncation() {
        let mut builder = FormulaParseDiagnosticsBuilder::new(FormulaParsePolicy::Warn);
        let formula = format!("={}", "A".repeat(119));
        assert_eq!(formula.len(), 120);

        builder.record_error("Sheet1", "A1", &formula, "unexpected token");
        let diagnostics = builder.build();

        let preview = &diagnostics.groups[0].formula_preview;
        assert!(preview.ends_with('…'));
        assert_ne!(preview, &formula);
        assert!(preview.len() <= FORMULA_PREVIEW_MAX_BYTES + '…'.len_utf8());
    }

    #[test]
    fn test_diagnostics_json_structure() {
        let mut builder = FormulaParseDiagnosticsBuilder::new(FormulaParsePolicy::Warn);
        builder.record_error("Sheet1", "A1", "=SUM(A:A)", "unexpected token");
        let diagnostics = builder.build();

        let value = serde_json::to_value(diagnostics).expect("serialize diagnostics");
        assert_eq!(value["policy"], serde_json::json!("warn"));
        assert_eq!(value["total_errors"], serde_json::json!(1));
        assert_eq!(value["groups_truncated"], serde_json::json!(false));
        assert!(value["groups"].is_array());

        let group = &value["groups"][0];
        assert_eq!(group["error_code"], serde_json::json!(FORMULA_PARSE_FAILED));
        assert_eq!(
            group["error_message"],
            serde_json::json!("unexpected token")
        );
        assert_eq!(group["sheet_name"], serde_json::json!("Sheet1"));
        assert_eq!(group["formula_preview"], serde_json::json!("=SUM(A:A)"));
        assert_eq!(group["count"], serde_json::json!(1));
        assert!(group["sample_addresses"].is_array());
    }

    #[test]
    fn test_has_errors_after_record() {
        let mut builder = FormulaParseDiagnosticsBuilder::new(FormulaParsePolicy::Warn);
        builder.record_error("Sheet1", "A1", "=SUM(A:A)", "unexpected token");

        assert!(builder.has_errors());
        assert!(!builder.is_empty());
    }
}
