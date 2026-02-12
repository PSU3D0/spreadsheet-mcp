use crate::cli::OutputFormat;
use anyhow::{bail, Result};
use serde::Serialize;

pub fn ensure_output_supported(format: OutputFormat) -> Result<()> {
    match format {
        OutputFormat::Json => Ok(()),
        OutputFormat::Csv => {
            bail!("csv output is not implemented yet for this CLI; use --format json")
        }
    }
}

#[derive(Debug, Serialize)]
pub struct ErrorEnvelope {
    pub code: String,
    pub message: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub did_you_mean: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub try_this: Option<String>,
}

pub fn envelope_for(error: &anyhow::Error) -> ErrorEnvelope {
    let message = error.to_string();

    if let Some((requested, suggested)) = parse_sheet_suggestion(&message) {
        return ErrorEnvelope {
            code: "SHEET_NOT_FOUND".to_string(),
            message: format!("sheet '{}' was not found", requested),
            did_you_mean: Some(suggested),
            try_this: Some(
                "run `agent-spreadsheet list-sheets <file>` to inspect valid names".to_string(),
            ),
        };
    }

    if message.contains("does not exist") {
        return ErrorEnvelope {
            code: "FILE_NOT_FOUND".to_string(),
            message,
            did_you_mean: None,
            try_this: Some("check the workbook path and permissions".to_string()),
        };
    }

    if message.contains("at least one range") {
        return ErrorEnvelope {
            code: "INVALID_ARGUMENT".to_string(),
            message,
            did_you_mean: None,
            try_this: Some("pass one or more A1 ranges, for example: `A1:C10`".to_string()),
        };
    }

    if message.contains("at least one edit") {
        return ErrorEnvelope {
            code: "INVALID_ARGUMENT".to_string(),
            message,
            did_you_mean: None,
            try_this: Some("add one or more edits like `A1=42` or `B2==SUM(A1:A1)`".to_string()),
        };
    }

    if message.contains("invalid shorthand edit") {
        return ErrorEnvelope {
            code: "INVALID_EDIT_SYNTAX".to_string(),
            message,
            did_you_mean: None,
            try_this: Some(
                "use `<cell>=<value>` for values or `<cell>==<formula>` for formulas".to_string(),
            ),
        };
    }

    if message.contains("csv output is not implemented") {
        return ErrorEnvelope {
            code: "OUTPUT_FORMAT_UNSUPPORTED".to_string(),
            message,
            did_you_mean: Some("json".to_string()),
            try_this: Some("re-run with `--format json`".to_string()),
        };
    }

    ErrorEnvelope {
        code: "COMMAND_FAILED".to_string(),
        message,
        did_you_mean: None,
        try_this: None,
    }
}

fn parse_sheet_suggestion(message: &str) -> Option<(String, String)> {
    let prefix = "sheet '";
    let not_found = "' not found; did you mean '";
    let suffix = "' ?";

    let start = message.find(prefix)? + prefix.len();
    let rest = &message[start..];
    let mid = rest.find(not_found)?;
    let requested = &rest[..mid];
    let suggestion_start = start + mid + not_found.len();
    let suggestion_rest = &message[suggestion_start..];
    let suggestion_end = suggestion_rest.find(suffix)?;
    let suggested = &suggestion_rest[..suggestion_end];
    Some((requested.to_string(), suggested.to_string()))
}
