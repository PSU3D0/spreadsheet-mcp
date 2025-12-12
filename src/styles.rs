use crate::model::{
    AlignmentDescriptor, BorderSideDescriptor, BordersDescriptor, FillDescriptor, FontDescriptor,
    GradientFillDescriptor, GradientStopDescriptor, PatternFillDescriptor, StyleDescriptor,
};
use sha2::{Digest, Sha256};
use std::collections::BTreeMap;
use umya_spreadsheet::structs::{EnumTrait, HorizontalAlignmentValues, VerticalAlignmentValues};
use umya_spreadsheet::{Border, Fill, Font, Style};

pub fn descriptor_from_style(style: &Style) -> StyleDescriptor {
    let font = style.get_font().and_then(descriptor_from_font);
    let fill = style.get_fill().and_then(descriptor_from_fill);
    let borders = style.get_borders().and_then(|borders| {
        let left = descriptor_from_border_side(borders.get_left_border());
        let right = descriptor_from_border_side(borders.get_right_border());
        let top = descriptor_from_border_side(borders.get_top_border());
        let bottom = descriptor_from_border_side(borders.get_bottom_border());
        let diagonal = descriptor_from_border_side(borders.get_diagonal_border());
        let vertical = descriptor_from_border_side(borders.get_vertical_border());
        let horizontal = descriptor_from_border_side(borders.get_horizontal_border());

        let diagonal_up = if *borders.get_diagonal_up() {
            Some(true)
        } else {
            None
        };
        let diagonal_down = if *borders.get_diagonal_down() {
            Some(true)
        } else {
            None
        };

        let descriptor = BordersDescriptor {
            left,
            right,
            top,
            bottom,
            diagonal,
            vertical,
            horizontal,
            diagonal_up,
            diagonal_down,
        };

        if descriptor.is_empty() {
            None
        } else {
            Some(descriptor)
        }
    });
    let alignment = style.get_alignment().and_then(descriptor_from_alignment);
    let number_format = style.get_number_format().and_then(|fmt| {
        let code = fmt.get_format_code();
        if code.eq_ignore_ascii_case("general") {
            None
        } else {
            Some(code.to_string())
        }
    });

    StyleDescriptor {
        font,
        fill,
        borders,
        alignment,
        number_format,
    }
}

pub fn stable_style_id(descriptor: &StyleDescriptor) -> String {
    let bytes = serde_json::to_vec(descriptor).unwrap_or_default();
    let mut hasher = Sha256::new();
    hasher.update(bytes);
    let digest = hasher.finalize();
    let hex = format!("{digest:x}");
    hex.chars().take(12).collect()
}

pub fn compress_positions_to_ranges(
    positions: &[(u32, u32)],
    limit: usize,
) -> (Vec<String>, bool) {
    if positions.is_empty() {
        return (Vec::new(), false);
    }

    let mut rows: BTreeMap<u32, Vec<u32>> = BTreeMap::new();
    for &(row, col) in positions {
        rows.entry(row).or_default().push(col);
    }
    for cols in rows.values_mut() {
        cols.sort_unstable();
        cols.dedup();
    }

    let mut spans_by_cols: BTreeMap<(u32, u32), Vec<u32>> = BTreeMap::new();
    for (row, cols) in rows {
        if cols.is_empty() {
            continue;
        }
        let mut start = cols[0];
        let mut prev = cols[0];
        for col in cols.into_iter().skip(1) {
            if col == prev + 1 {
                prev = col;
            } else {
                spans_by_cols.entry((start, prev)).or_default().push(row);
                start = col;
                prev = col;
            }
        }
        spans_by_cols.entry((start, prev)).or_default().push(row);
    }

    let mut ranges = Vec::new();
    let mut truncated = false;

    'outer: for ((start_col, end_col), mut span_rows) in spans_by_cols {
        span_rows.sort_unstable();
        span_rows.dedup();
        if span_rows.is_empty() {
            continue;
        }
        let mut run_start = span_rows[0];
        let mut prev_row = span_rows[0];
        for row in span_rows.into_iter().skip(1) {
            if row == prev_row + 1 {
                prev_row = row;
                continue;
            }
            ranges.push(format_range(start_col, end_col, run_start, prev_row));
            if ranges.len() >= limit {
                truncated = true;
                break 'outer;
            }
            run_start = row;
            prev_row = row;
        }
        ranges.push(format_range(start_col, end_col, run_start, prev_row));
        if ranges.len() >= limit {
            truncated = true;
            break;
        }
    }

    if truncated {
        ranges.truncate(limit);
    }
    (ranges, truncated)
}

fn format_range(start_col: u32, end_col: u32, start_row: u32, end_row: u32) -> String {
    let start_addr = crate::utils::cell_address(start_col, start_row);
    let end_addr = crate::utils::cell_address(end_col, end_row);
    if start_addr == end_addr {
        start_addr
    } else {
        format!("{start_addr}:{end_addr}")
    }
}

fn descriptor_from_font(font: &Font) -> Option<FontDescriptor> {
    let bold = *font.get_bold();
    let italic = *font.get_italic();
    let underline = font.get_underline();
    let strikethrough = *font.get_strikethrough();
    let color = font.get_color().get_argb();

    let descriptor = FontDescriptor {
        name: Some(font.get_name().to_string()).filter(|s| !s.is_empty()),
        size: Some(*font.get_size()).filter(|s| *s > 0.0),
        bold: if bold { Some(true) } else { None },
        italic: if italic { Some(true) } else { None },
        underline: if underline.eq_ignore_ascii_case("none") {
            None
        } else {
            Some(underline.to_string())
        },
        strikethrough: if strikethrough { Some(true) } else { None },
        color: Some(color.to_string()).filter(|s| !s.is_empty()),
    };

    if descriptor.is_empty() {
        None
    } else {
        Some(descriptor)
    }
}

fn descriptor_from_fill(fill: &Fill) -> Option<FillDescriptor> {
    if let Some(pattern) = fill.get_pattern_fill() {
        let pattern_type = pattern.get_pattern_type();
        let kind = pattern_type.get_value_string();
        let fg = pattern
            .get_foreground_color()
            .map(|c| c.get_argb().to_string())
            .filter(|s| !s.is_empty());
        let bg = pattern
            .get_background_color()
            .map(|c| c.get_argb().to_string())
            .filter(|s| !s.is_empty());

        if kind.eq_ignore_ascii_case("none") && fg.is_none() && bg.is_none() {
            return None;
        }

        return Some(FillDescriptor::Pattern(PatternFillDescriptor {
            pattern_type: if kind.eq_ignore_ascii_case("none") {
                None
            } else {
                Some(kind.to_string())
            },
            foreground_color: fg,
            background_color: bg,
        }));
    }

    if let Some(gradient) = fill.get_gradient_fill() {
        let stops: Vec<GradientStopDescriptor> = gradient
            .get_gradient_stop()
            .iter()
            .map(|stop| GradientStopDescriptor {
                position: *stop.get_position(),
                color: stop.get_color().get_argb().to_string(),
            })
            .collect();

        let degree = *gradient.get_degree();
        if stops.is_empty() && degree == 0.0 {
            return None;
        }

        return Some(FillDescriptor::Gradient(GradientFillDescriptor {
            degree: Some(degree).filter(|d| *d != 0.0),
            stops,
        }));
    }

    None
}

fn descriptor_from_border_side(border: &Border) -> Option<BorderSideDescriptor> {
    let style = border.get_border_style();
    let style = if style.eq_ignore_ascii_case("none") {
        None
    } else {
        Some(style.to_string())
    };
    let color = Some(border.get_color().get_argb().to_string()).filter(|s| !s.is_empty());

    let descriptor = BorderSideDescriptor { style, color };
    if descriptor.is_empty() {
        None
    } else {
        Some(descriptor)
    }
}

fn descriptor_from_alignment(alignment: &umya_spreadsheet::Alignment) -> Option<AlignmentDescriptor> {
    let horizontal = if alignment.get_horizontal() != &HorizontalAlignmentValues::General {
        Some(alignment.get_horizontal().get_value_string().to_string())
    } else {
        None
    };
    let vertical = if alignment.get_vertical() != &VerticalAlignmentValues::Bottom {
        Some(alignment.get_vertical().get_value_string().to_string())
    } else {
        None
    };
    let wrap_text = if *alignment.get_wrap_text() {
        Some(true)
    } else {
        None
    };
    let text_rotation = if *alignment.get_text_rotation() != 0 {
        Some(*alignment.get_text_rotation())
    } else {
        None
    };

    let descriptor = AlignmentDescriptor {
        horizontal,
        vertical,
        wrap_text,
        text_rotation,
    };
    if descriptor.is_empty() {
        None
    } else {
        Some(descriptor)
    }
}

trait IsEmpty {
    fn is_empty(&self) -> bool;
}

impl IsEmpty for FontDescriptor {
    fn is_empty(&self) -> bool {
        self.name.is_none()
            && self.size.is_none()
            && self.bold.is_none()
            && self.italic.is_none()
            && self.underline.is_none()
            && self.strikethrough.is_none()
            && self.color.is_none()
    }
}

impl IsEmpty for BorderSideDescriptor {
    fn is_empty(&self) -> bool {
        self.style.is_none() && self.color.is_none()
    }
}

impl IsEmpty for BordersDescriptor {
    fn is_empty(&self) -> bool {
        self.left.is_none()
            && self.right.is_none()
            && self.top.is_none()
            && self.bottom.is_none()
            && self.diagonal.is_none()
            && self.vertical.is_none()
            && self.horizontal.is_none()
            && self.diagonal_up.is_none()
            && self.diagonal_down.is_none()
    }
}

impl IsEmpty for AlignmentDescriptor {
    fn is_empty(&self) -> bool {
        self.horizontal.is_none()
            && self.vertical.is_none()
            && self.wrap_text.is_none()
            && self.text_rotation.is_none()
    }
}
