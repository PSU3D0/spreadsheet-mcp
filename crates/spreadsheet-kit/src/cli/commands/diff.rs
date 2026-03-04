use crate::runtime::stateless::StatelessRuntime;
use anyhow::{Result, anyhow, bail};
use serde_json::{Map, Value, json};
use std::collections::{BTreeMap, BTreeSet};
use std::path::PathBuf;

const DIFF_LIMIT_MAX: u32 = 2_000;

#[derive(Debug, Clone, Copy)]
struct A1Bounds {
    start_col: u32,
    end_col: u32,
    start_row: u32,
    end_row: u32,
}

pub async fn diff(
    original: PathBuf,
    modified: PathBuf,
    sheet: Option<String>,
    sheets: Option<Vec<String>>,
    range: Option<String>,
    details: bool,
    limit: u32,
    offset: u32,
) -> Result<Value> {
    if sheet.is_some() && sheets.is_some() {
        bail!("invalid argument: --sheet and --sheets are mutually exclusive");
    }

    let runtime = StatelessRuntime;
    let original = runtime.normalize_existing_file(&original)?;
    let modified = runtime.normalize_existing_file(&modified)?;

    if details && (limit == 0 || limit > DIFF_LIMIT_MAX) {
        bail!(
            "invalid argument: --limit must be between 1 and {}",
            DIFF_LIMIT_MAX
        );
    }

    let sheet_filters: Vec<String> = if let Some(s) = sheet {
        vec![s]
    } else {
        sheets.unwrap_or_default()
    };

    let range_bounds = if let Some(range) = range.as_ref() {
        Some(
            parse_a1_range(range)
                .ok_or_else(|| anyhow!("invalid argument: --range must be A1 notation"))?,
        )
    } else {
        None
    };

    let mut payload = runtime.diff_json(&original, &modified)?;
    let changes = payload
        .get_mut("changes")
        .and_then(Value::as_array_mut)
        .map(std::mem::take)
        .unwrap_or_default();

    let mut counts_by_kind: BTreeMap<String, u32> = BTreeMap::new();
    let mut counts_by_type: BTreeMap<String, u32> = BTreeMap::new();
    let mut affected_sheets: BTreeSet<String> = BTreeSet::new();

    let mut filtered = Vec::new();
    for change in changes {
        if !change_matches_filters(&change, &sheet_filters, range_bounds) {
            continue;
        }

        let kind = change_kind(&change).to_string();
        *counts_by_kind.entry(kind).or_default() += 1;

        let type_key = change
            .get("type")
            .and_then(Value::as_str)
            .unwrap_or("unknown")
            .to_string();
        *counts_by_type.entry(type_key).or_default() += 1;

        if let Some(sheet_name) = change_sheet_name(&change) {
            affected_sheets.insert(sheet_name.to_string());
        }

        filtered.push(change);
    }

    let total_changes = filtered.len() as u32;

    let (returned_changes, paged_changes, truncated, next_offset) = if details {
        let offset = offset as usize;
        let limit = limit as usize;
        let total = filtered.len();
        let page: Vec<Value> = filtered.into_iter().skip(offset).take(limit).collect();
        let returned = page.len() as u32;
        let consumed = offset.saturating_add(returned as usize);
        let truncated = consumed < total;
        let next_offset = truncated.then_some(consumed as u32);
        (returned, page, truncated, next_offset)
    } else {
        (0, Vec::new(), false, None)
    };

    let summary = json!({
        "total_changes": total_changes,
        "returned_changes": returned_changes,
        "truncated": truncated,
        "next_offset": next_offset,
        "counts_by_kind": counts_by_kind,
        "counts_by_type": counts_by_type,
        "affected_sheets": affected_sheets.into_iter().collect::<Vec<_>>()
    });

    let mut response = Map::new();
    response.insert(
        "original".to_string(),
        Value::String(original.display().to_string()),
    );
    response.insert(
        "modified".to_string(),
        Value::String(modified.display().to_string()),
    );
    response.insert("change_count".to_string(), Value::from(total_changes));
    response.insert("summary".to_string(), summary);

    if details {
        response.insert("changes".to_string(), Value::Array(paged_changes));
    }

    Ok(Value::Object(response))
}

fn change_kind(change: &Value) -> &'static str {
    if change.get("address").is_some() {
        "cell"
    } else if change.get("display_name").is_some() {
        "table"
    } else if change.get("name").is_some() {
        "name"
    } else {
        "unknown"
    }
}

fn change_sheet_name(change: &Value) -> Option<&str> {
    change
        .get("sheet")
        .and_then(Value::as_str)
        .or_else(|| change.get("scope_sheet").and_then(Value::as_str))
}

fn change_matches_filters(
    change: &Value,
    sheet_filters: &[String],
    range: Option<A1Bounds>,
) -> bool {
    if !sheet_filters.is_empty() {
        let Some(sheet_name) = change_sheet_name(change) else {
            return false;
        };
        if !sheet_filters
            .iter()
            .any(|f| sheet_name.eq_ignore_ascii_case(f))
        {
            return false;
        }
    }

    let Some(bounds) = range else {
        return true;
    };

    if let Some(address) = change.get("address").and_then(Value::as_str) {
        return address_in_bounds(address, bounds);
    }

    ["range", "old_range", "new_range"]
        .iter()
        .filter_map(|key| change.get(*key).and_then(Value::as_str))
        .any(|candidate| range_intersects(candidate, bounds))
}

fn address_in_bounds(address: &str, bounds: A1Bounds) -> bool {
    let Some((col, row)) = parse_a1_coord(address) else {
        return false;
    };
    col >= bounds.start_col
        && col <= bounds.end_col
        && row >= bounds.start_row
        && row <= bounds.end_row
}

fn range_intersects(range: &str, bounds: A1Bounds) -> bool {
    let Some(candidate) = parse_a1_range(range) else {
        return false;
    };

    !(candidate.end_col < bounds.start_col
        || candidate.start_col > bounds.end_col
        || candidate.end_row < bounds.start_row
        || candidate.start_row > bounds.end_row)
}

fn parse_a1_range(raw: &str) -> Option<A1Bounds> {
    let mut text = raw.trim();
    if text.is_empty() {
        return None;
    }

    if let Some((_, tail)) = text.rsplit_once('!') {
        text = tail;
    }

    let (left, right) = text.split_once(':').map_or((text, text), |(a, b)| (a, b));
    let (c1, r1) = parse_a1_coord(left)?;
    let (c2, r2) = parse_a1_coord(right)?;

    Some(A1Bounds {
        start_col: c1.min(c2),
        end_col: c1.max(c2),
        start_row: r1.min(r2),
        end_row: r1.max(r2),
    })
}

fn parse_a1_coord(raw: &str) -> Option<(u32, u32)> {
    let coord = raw.trim().trim_start_matches('$');
    if coord.is_empty() {
        return None;
    }

    let mut letters = String::new();
    let mut digits = String::new();
    for ch in coord.chars() {
        if ch == '$' {
            continue;
        }
        if ch.is_ascii_alphabetic() {
            if !digits.is_empty() {
                return None;
            }
            letters.push(ch.to_ascii_uppercase());
        } else if ch.is_ascii_digit() {
            digits.push(ch);
        } else {
            return None;
        }
    }

    if letters.is_empty() || digits.is_empty() {
        return None;
    }

    let mut col = 0u32;
    for ch in letters.bytes() {
        col = col
            .saturating_mul(26)
            .saturating_add((ch - b'A' + 1) as u32);
    }

    let row: u32 = digits.parse().ok()?;
    if col == 0 || row == 0 {
        return None;
    }

    Some((col, row))
}
