use crate::cli::{OutputFormat, OutputShape};
use crate::response_prune::prune_non_structural_empties;
use anyhow::{Result, bail};
use serde_json::{Map, Value};

pub fn emit_value(
    value: &Value,
    format: OutputFormat,
    shape: OutputShape,
    compact: bool,
    quiet: bool,
) -> Result<()> {
    if matches!(format, OutputFormat::Csv) {
        bail!("csv output is not implemented yet for agent-spreadsheet")
    }

    let mut value = value.clone();
    prune_non_structural_empties(&mut value);
    apply_shape(&mut value, shape);

    let stdout = std::io::stdout();
    let mut handle = stdout.lock();
    if compact || quiet {
        serde_json::to_writer(&mut handle, &value)?;
    } else {
        serde_json::to_writer_pretty(&mut handle, &value)?;
    }
    use std::io::Write;
    handle.write_all(b"\n")?;
    Ok(())
}

fn apply_shape(value: &mut Value, shape: OutputShape) {
    if !matches!(shape, OutputShape::Compact) {
        return;
    }

    remove_workbook_short_id(value);
    flatten_single_range_values(value);
}

fn remove_workbook_short_id(value: &mut Value) {
    match value {
        Value::Object(obj) => {
            obj.remove("workbook_short_id");
            for nested in obj.values_mut() {
                remove_workbook_short_id(nested);
            }
        }
        Value::Array(items) => {
            for item in items {
                remove_workbook_short_id(item);
            }
        }
        _ => {}
    }
}

fn flatten_single_range_values(value: &mut Value) {
    let Value::Object(obj) = value else {
        return;
    };

    let Some(Value::Array(entries)) = obj.get("values") else {
        return;
    };

    if entries.len() != 1 {
        return;
    }

    let Some(Value::Object(entry)) = entries.first() else {
        return;
    };

    let looks_like_range_values_entry = entry.get("range").is_some()
        && (entry.contains_key("rows")
            || entry.contains_key("values")
            || entry.contains_key("csv"));
    if !looks_like_range_values_entry {
        return;
    }

    let entry_fields = entry.clone();
    obj.remove("values");
    merge_entry_fields(obj, entry_fields);
}

fn merge_entry_fields(target: &mut Map<String, Value>, entry_fields: Map<String, Value>) {
    for (key, value) in entry_fields {
        target.insert(key, value);
    }
}
