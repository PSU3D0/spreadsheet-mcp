use crate::utils::column_number_to_name;
use anyhow::{Result, anyhow, bail};
use formualizer_parse::parser::ReferenceType;
use formualizer_parse::{ASTNode, ASTNodeType, LiteralValue};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RelativeMode {
    Excel,
    AbsCols,
    AbsRows,
}

impl RelativeMode {
    pub fn parse(mode: Option<&str>) -> Result<Self> {
        match mode.unwrap_or("excel").to_ascii_lowercase().as_str() {
            "excel" => Ok(Self::Excel),
            "abs_cols" | "abscols" | "columns_absolute" => Ok(Self::AbsCols),
            "abs_rows" | "absrows" | "rows_absolute" => Ok(Self::AbsRows),
            other => bail!("invalid relative_mode: {}", other),
        }
    }
}

#[derive(Debug, Clone, Copy)]
struct CoordFlags {
    abs_col: bool,
    abs_row: bool,
}

pub fn parse_base_formula(formula: &str) -> Result<ASTNode> {
    let trimmed = formula.trim();
    let with_equals = if trimmed.starts_with('=') {
        trimmed.to_string()
    } else {
        format!("={}", trimmed)
    };
    formualizer_parse::parse(&with_equals)
        .map_err(|e| anyhow!("failed to parse base_formula: {}", e.message))
}

pub fn shift_formula_ast(
    ast: &ASTNode,
    delta_col: i32,
    delta_row: i32,
    mode: RelativeMode,
) -> Result<String> {
    Ok(format!("={}", shift_node(ast, delta_col, delta_row, mode)?))
}

fn shift_node(
    node: &ASTNode,
    delta_col: i32,
    delta_row: i32,
    mode: RelativeMode,
) -> Result<String> {
    Ok(match &node.node_type {
        ASTNodeType::Literal(value) => match value {
            LiteralValue::Text(s) => {
                let escaped = s.replace('"', "\"\"");
                format!("\"{escaped}\"")
            }
            _ => format!("{value}"),
        },
        ASTNodeType::Reference {
            original,
            reference,
        } => shift_reference(original, reference, delta_col, delta_row, mode)?,
        ASTNodeType::UnaryOp { op, expr } => {
            format!("{}{}", op, shift_node(expr, delta_col, delta_row, mode)?)
        }
        ASTNodeType::BinaryOp { op, left, right } => {
            if op == ":" {
                format!(
                    "{}:{}",
                    shift_node(left, delta_col, delta_row, mode)?,
                    shift_node(right, delta_col, delta_row, mode)?
                )
            } else {
                format!(
                    "{} {} {}",
                    shift_node(left, delta_col, delta_row, mode)?,
                    op,
                    shift_node(right, delta_col, delta_row, mode)?
                )
            }
        }
        ASTNodeType::Function { name, args } => {
            let args_str = args
                .iter()
                .map(|a| shift_node(a, delta_col, delta_row, mode))
                .collect::<Result<Vec<_>>>()?
                .join(", ");
            format!("{}({})", name.to_uppercase(), args_str)
        }
        ASTNodeType::Array(rows) => {
            let rows_str = rows
                .iter()
                .map(|row| {
                    row.iter()
                        .map(|a| shift_node(a, delta_col, delta_row, mode))
                        .collect::<Result<Vec<_>>>()
                        .map(|parts| parts.join(", "))
                })
                .collect::<Result<Vec<_>>>()?
                .join("; ");
            format!("{{{rows_str}}}")
        }
    })
}

fn shift_reference(
    original: &str,
    reference: &ReferenceType,
    delta_col: i32,
    delta_row: i32,
    mode: RelativeMode,
) -> Result<String> {
    match reference {
        ReferenceType::Cell { sheet, row, col } => {
            let coord_part = strip_sheet_prefix(original);
            let mut flags = coord_abs_flags(coord_part);
            match mode {
                RelativeMode::AbsCols => flags.abs_col = true,
                RelativeMode::AbsRows => flags.abs_row = true,
                RelativeMode::Excel => {}
            }
            let new_col = shift_u32(*col, flags.abs_col, delta_col)?;
            let new_row = shift_u32(*row, flags.abs_row, delta_row)?;
            let coord = format_cell_coord(new_col, new_row, flags);
            Ok(format!("{}{}", format_sheet_prefix(sheet), coord))
        }
        ReferenceType::Range {
            sheet,
            start_row,
            start_col,
            end_row,
            end_col,
        } => {
            let ref_part = strip_sheet_prefix(original);
            let (start_str, end_str) = ref_part.split_once(':').unwrap_or((ref_part, ref_part));
            let mut start_flags = coord_abs_flags(start_str);
            let mut end_flags = coord_abs_flags(end_str);

            match mode {
                RelativeMode::AbsCols => {
                    if start_col.is_some() {
                        start_flags.abs_col = true;
                    }
                    if end_col.is_some() {
                        end_flags.abs_col = true;
                    }
                }
                RelativeMode::AbsRows => {
                    if start_row.is_some() {
                        start_flags.abs_row = true;
                    }
                    if end_row.is_some() {
                        end_flags.abs_row = true;
                    }
                }
                RelativeMode::Excel => {}
            }

            let new_start_col = shift_opt_u32(*start_col, start_flags.abs_col, delta_col)?;
            let new_end_col = shift_opt_u32(*end_col, end_flags.abs_col, delta_col)?;
            let new_start_row = shift_opt_u32(*start_row, start_flags.abs_row, delta_row)?;
            let new_end_row = shift_opt_u32(*end_row, end_flags.abs_row, delta_row)?;

            let start_coord = format_range_coord(new_start_col, new_start_row, start_flags);
            let end_coord = format_range_coord(new_end_col, new_end_row, end_flags);
            if start_coord.is_empty() || end_coord.is_empty() {
                bail!("invalid range reference after shift: {}", original);
            }
            let coord = format!("{start_coord}:{end_coord}");
            Ok(format!("{}{}", format_sheet_prefix(sheet), coord))
        }
        ReferenceType::Table(_) | ReferenceType::NamedRange(_) => Ok(reference.to_string()),
    }
}

fn shift_u32(value: u32, abs: bool, delta: i32) -> Result<u32> {
    if abs || delta == 0 {
        return Ok(value);
    }
    let shifted = value as i64 + delta as i64;
    if shifted < 1 {
        bail!("shift would move reference before A1");
    }
    Ok(shifted as u32)
}

fn shift_opt_u32(value: Option<u32>, abs: bool, delta: i32) -> Result<Option<u32>> {
    match value {
        Some(v) => Ok(Some(shift_u32(v, abs, delta)?)),
        None => Ok(None),
    }
}

fn strip_sheet_prefix(original: &str) -> &str {
    original
        .rsplit_once('!')
        .map(|(_, tail)| tail)
        .unwrap_or(original)
        .trim()
}

fn coord_abs_flags(coord: &str) -> CoordFlags {
    let bytes = coord.as_bytes();
    let len = bytes.len();
    let mut i = 0;
    let leading_dollar = i < len && bytes[i] == b'$';
    if leading_dollar {
        i += 1;
    }

    let letters_start = i;
    while i < len && bytes[i].is_ascii_alphabetic() {
        i += 1;
    }
    let has_letters = i > letters_start;

    let second_dollar = i < len && bytes[i] == b'$';
    let digits_start = if second_dollar { i + 1 } else { i };
    let mut j = digits_start;
    while j < len && bytes[j].is_ascii_digit() {
        j += 1;
    }
    let has_digits = j > digits_start;

    let abs_col = leading_dollar && has_letters;
    let abs_row = if has_letters {
        second_dollar && has_digits
    } else {
        leading_dollar && has_digits
    };

    CoordFlags { abs_col, abs_row }
}

fn format_cell_coord(col: u32, row: u32, flags: CoordFlags) -> String {
    let col_str = column_number_to_name(col);
    let mut out = String::new();
    if flags.abs_col {
        out.push('$');
    }
    out.push_str(&col_str);
    if flags.abs_row {
        out.push('$');
    }
    out.push_str(&row.to_string());
    out
}

fn format_range_coord(col: Option<u32>, row: Option<u32>, flags: CoordFlags) -> String {
    match (col, row) {
        (Some(c), Some(r)) => format_cell_coord(c, r, flags),
        (Some(c), None) => {
            let col_str = column_number_to_name(c);
            if flags.abs_col {
                format!("${col_str}")
            } else {
                col_str
            }
        }
        (None, Some(r)) => {
            if flags.abs_row {
                format!("${r}")
            } else {
                r.to_string()
            }
        }
        (None, None) => String::new(),
    }
}

fn format_sheet_prefix(sheet: &Option<String>) -> String {
    if let Some(name) = sheet {
        if sheet_name_needs_quoting(name) {
            let escaped = name.replace('\'', "''");
            format!("'{escaped}'!")
        } else {
            format!("{name}!")
        }
    } else {
        String::new()
    }
}

fn sheet_name_needs_quoting(name: &str) -> bool {
    if name.is_empty() {
        return false;
    }
    let bytes = name.as_bytes();
    if bytes[0].is_ascii_digit() {
        return true;
    }
    for &byte in bytes {
        match byte {
            b' ' | b'!' | b'"' | b'#' | b'$' | b'%' | b'&' | b'\'' | b'(' | b')' | b'*' | b'+'
            | b',' | b'-' | b'.' | b'/' | b':' | b';' | b'<' | b'=' | b'>' | b'?' | b'@' | b'['
            | b'\\' | b']' | b'^' | b'`' | b'{' | b'|' | b'}' | b'~' => return true,
            _ => {}
        }
    }
    let upper = name.to_uppercase();
    matches!(
        upper.as_str(),
        "TRUE" | "FALSE" | "NULL" | "REF" | "DIV" | "NAME" | "NUM" | "VALUE" | "N/A"
    )
}
