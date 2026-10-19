// Text-focused markdown projection for xlsx workbooks. Walks a parsed
// Worksheet JSON value and emits a `## SheetName` heading followed by a
// pipe table containing the cells' cached display values. Designed for AI
// agents that need to read spreadsheet content efficiently — drops
// styling, formatting (numFmt), charts, sparklines, drawings, slicers,
// conditional formatting, and the formula text (cached value only).

use std::borrow::Cow;
use std::collections::{BTreeMap, HashMap};
use std::fmt::Write as _;

use ooxml_common::markdown::{escape_text, split_lines, TextContext};
use serde_json::Value;

use crate::types::SharedString;

/// Append one sheet's table to `out`. `Err(projected)` reports that the
/// projected output would exceed `limit` bytes; the caller then fails the
/// whole projection, so the partial `out` is discarded.
pub(crate) fn render_sheet(
    sheet: &Value,
    shared_strings: &[SharedString],
    out: &mut String,
    limit: u64,
) -> Result<(), u64> {
    let name = sheet["name"].as_str().unwrap_or("(unnamed)");
    let _ = writeln!(out, "## {}\n", escape_text(name, TextContext::Heading));
    if out.len() as u64 > limit {
        return Err(out.len() as u64);
    }

    let Some(rows) = sheet["rows"].as_array() else {
        return Ok(());
    };

    // Find the populated bounding box. Rows are stored sparsely keyed by
    // 1-based row index; cells likewise carry their 1-based col. The table
    // spans [min_row..=max_row] × [min_col..=max_col], with gaps rendered as
    // empty cells. Cells are kept sparse: the bounding box of two cells at A1
    // and XFD1048576 must not allocate a dense 1048576 × 16384 grid.
    let mut min_row = u32::MAX;
    let mut max_row = 0u32;
    let mut min_col = u32::MAX;
    let mut max_col = 0u32;
    for row in rows {
        let row_idx = row["index"].as_u64().unwrap_or(0) as u32;
        let Some(cells) = row["cells"].as_array() else {
            continue;
        };
        for cell in cells {
            let s = cell_display(cell, shared_strings);
            if s.is_empty() {
                continue;
            }
            let col = cell["col"].as_u64().unwrap_or(0) as u32;
            if row_idx == 0 || col == 0 {
                continue;
            }
            min_row = min_row.min(row_idx);
            max_row = max_row.max(row_idx);
            min_col = min_col.min(col);
            max_col = max_col.max(col);
        }
    }
    if max_row == 0 || max_col == 0 {
        // Empty sheet — emit the heading only.
        return Ok(());
    }

    // Sparse grid of display strings keyed by (row, col) inside the bbox.
    // Display strings borrow from the sheet JSON and the shared-string table,
    // so repeated references to one long string cost nothing until a row is
    // written, which happens only after the size admission below.
    let mut grid: BTreeMap<(u32, u32), Cow<'_, str>> = BTreeMap::new();
    for row in rows {
        let row_idx = row["index"].as_u64().unwrap_or(0) as u32;
        if row_idx < min_row || row_idx > max_row {
            continue;
        }
        let Some(cells) = row["cells"].as_array() else {
            continue;
        };
        for cell in cells {
            let col = cell["col"].as_u64().unwrap_or(0) as u32;
            if col < min_col || col > max_col {
                continue;
            }
            grid.insert((row_idx, col), cell_display(cell, shared_strings));
        }
    }

    // Apply merges: ECMA-376 §18.3.1.55 — the top-left cell carries the value,
    // continuation cells must render empty. We do this after grid population so
    // a value living at the top-left of a merge survives and the rest clear.
    clear_merge_continuation_cells(&sheet["mergeCells"], &mut grid);

    // Every emitted row carries all `n_cols` cells, so bound the projected
    // output before writing it: a few far-apart cells can otherwise expand
    // into gigabytes of empty `|  |` columns. A row costs at most 6 bytes per
    // column (the `--- | ` separator) plus its delimiters and newline.
    let n_cols = u64::from(max_col - min_col + 1);
    let mut emitted_rows = 2u64; // header + separator
    let mut content = 0u64;
    let mut previous_row = None;
    for (&(row_idx, _), value) in &grid {
        content = content.saturating_add(value.len() as u64 * 4);
        if row_idx != min_row && !value.is_empty() && previous_row != Some(row_idx) {
            emitted_rows += 1;
            previous_row = Some(row_idx);
        }
    }
    let projected = (out.len() as u64)
        .saturating_add(emitted_rows.saturating_mul(n_cols.saturating_mul(6).saturating_add(4)))
        .saturating_add(content)
        .saturating_add(1);
    if projected > limit {
        return Err(projected);
    }

    let n_cols = n_cols as usize;
    let row_cells = |row_idx: u32| -> Vec<&str> {
        let mut cells = vec![""; n_cols];
        for ((_, col), value) in grid.range((row_idx, min_col)..=(row_idx, max_col)) {
            cells[(col - min_col) as usize] = value.as_ref();
        }
        cells
    };
    // Header row: use the first row of the bbox. Markdown tables require a
    // header — if the first row is blank we still emit empty headers so
    // downstream renderers parse the rest correctly.
    write_table_row(out, &row_cells(min_row), n_cols);
    let sep: Vec<&str> = (0..n_cols).map(|_| "---").collect();
    let _ = writeln!(out, "| {} |", sep.join(" | "));
    let mut written_row = min_row;
    for (&(row_idx, _), value) in &grid {
        // Skip fully-empty middle rows to keep output tight.
        if row_idx == written_row || value.is_empty() {
            continue;
        }
        written_row = row_idx;
        write_table_row(out, &row_cells(row_idx), n_cols);
    }
    out.push('\n');
    Ok(())
}

fn write_table_row(out: &mut String, row: &[&str], n_cols: usize) {
    let cells: Vec<String> = (0..n_cols)
        .map(|i| {
            row.get(i)
                .map(|s| escape_text(s, TextContext::TableCell))
                .unwrap_or_default()
        })
        .collect();
    let _ = writeln!(out, "| {} |", cells.join(" | "));
}

fn cell_display<'a>(cell: &'a Value, shared_strings: &'a [SharedString]) -> Cow<'a, str> {
    // CellValue has `rename_all = "camelCase"` so the JSON tag is lowercase
    // ("text"/"number"/...). PascalCase would silently never match — same
    // class of bug that hid pptx_extract_text earlier.
    let value = &cell["value"];
    match value["type"].as_str().unwrap_or("empty") {
        "text" => Cow::Borrowed(value["text"].as_str().unwrap_or("")),
        // A `t="s"` cell ships only an `si` index now; resolve it back to the
        // shared-string table's plain text (markdown drops runs anyway,
        // matching the `"text"` arm).
        "shared" => value["si"]
            .as_u64()
            .and_then(|i| shared_strings.get(i as usize))
            .map(|s| Cow::Borrowed(s.text.as_str()))
            .unwrap_or_default(),
        "number" => value["number"]
            .as_f64()
            .map(|n| Cow::Owned(format_number(n)))
            .unwrap_or_default(),
        "bool" => value["bool"]
            .as_bool()
            .map(|b| Cow::Borrowed(if b { "TRUE" } else { "FALSE" }))
            .unwrap_or_default(),
        "error" => Cow::Borrowed(value["error"].as_str().unwrap_or("#ERR")),
        _ => Cow::Borrowed(""),
    }
}

fn format_number(n: f64) -> String {
    // Integer-valued doubles → integer form so 2025 doesn't show as 2025.0.
    if n.is_finite() && n.fract() == 0.0 && n.abs() < 1e15 {
        return format!("{}", n as i64);
    }
    // Round to 10 significant digits to mask IEEE-754 ULP noise (702.6
    // round-trips through XML as 702.5999999999999). Trim trailing zeros so
    // 702.6 doesn't render as 702.6000000000.
    let s = format!("{n:.10}");
    let trimmed = s.trim_end_matches('0').trim_end_matches('.').to_string();
    if trimmed.is_empty() {
        "0".to_string()
    } else {
        trimmed
    }
}

/// SpreadsheetML grid maxima (ECMA-376 §18.3.1.4 / Excel's 16384 × 1048576).
const MAX_COL: u32 = 16_384;
const MAX_ROW: u32 = 1_048_576;

/// Clear every populated cell that is a continuation cell of a merged range —
/// any cell of `[top..=bottom] × [left..=right]` except that range's top-left.
///
/// A merge may cover the whole sheet, so its cells are never enumerated.
/// Instead a row sweep keeps, in a Fenwick tree over columns, how many merges
/// cover each column of the current row; a populated cell is a continuation
/// cell exactly when more merges cover it than start at it. The work is
/// O((cells + merges) · log columns) regardless of merge area.
fn clear_merge_continuation_cells(
    merge_cells: &Value,
    grid: &mut BTreeMap<(u32, u32), Cow<'_, str>>,
) {
    let Some(arr) = merge_cells.as_array() else {
        return;
    };
    // (row, column-range add): +1 at `top`, -1 after `bottom`.
    let mut events: Vec<(u32, u32, u32, i64)> = Vec::new();
    let mut starts: HashMap<(u32, u32), i64> = HashMap::new();
    for m in arr {
        let coord = |key: &str| {
            m[key]
                .as_u64()
                .and_then(|v| u32::try_from(v).ok())
                .unwrap_or(0)
        };
        let (top, left, bottom, right) =
            (coord("top"), coord("left"), coord("bottom"), coord("right"));
        if top == 0
            || left == 0
            || bottom < top
            || right < left
            || bottom > MAX_ROW
            || right > MAX_COL
        {
            continue;
        }
        events.push((top, left, right, 1));
        events.push((bottom + 1, left, right, -1));
        *starts.entry((top, left)).or_default() += 1;
    }
    if events.is_empty() {
        return;
    }
    events.sort_unstable_by_key(|event| event.0);
    let mut tree = vec![0i64; MAX_COL as usize + 2];
    // Fenwick tree over 1-based columns: an update climbs by the lowest set
    // bit (`1 << trailing_zeros`), a prefix query drops it (`i &= i - 1`).
    let add = |tree: &mut [i64], index: u32, delta: i64| {
        let mut i = index as usize;
        while i > 0 && i < tree.len() {
            tree[i] += delta;
            i += 1 << i.trailing_zeros();
        }
    };
    let covering = |tree: &[i64], index: u32| {
        let (mut i, mut sum) = (index as usize, 0i64);
        while i > 0 {
            sum += tree[i];
            i &= i - 1;
        }
        sum
    };
    let mut next = 0;
    for (&(row, col), value) in grid.iter_mut() {
        while next < events.len() && events[next].0 <= row {
            let (_, left, right, delta) = events[next];
            add(&mut tree, left, delta);
            add(&mut tree, right + 1, -delta);
            next += 1;
        }
        let starting_here = starts.get(&(row, col)).copied().unwrap_or(0);
        if covering(&tree, col) > starting_here {
            *value = Cow::Borrowed("");
        }
    }
}

/// Append one sheet's review comments to `out`. `Err(len)` reports that `out`
/// grew past `limit` bytes; it is checked after every comment and reply, so
/// repeated per-comment headings (each carrying the sheet name) cannot grow
/// the buffer by more than one entry beyond the budget.
pub(crate) fn render_sheet_comments(
    sheet: &Value,
    out: &mut String,
    has_comments: &mut bool,
    limit: u64,
) -> Result<(), u64> {
    let over = |out: &String| (out.len() as u64 > limit).then_some(out.len() as u64);
    let Some(comments) = sheet["comments"].as_array() else {
        return Ok(());
    };
    if comments.is_empty() {
        return Ok(());
    }
    if !*has_comments {
        out.push_str("\n## Review comments\n\n");
        *has_comments = true;
    }
    let sheet_name = sheet["name"].as_str().unwrap_or("(unnamed)");
    for comment in comments {
        let cell = comment["cellRef"].as_str().unwrap_or("(unknown cell)");
        let status = if comment["resolved"].as_bool() == Some(true) {
            " (resolved)"
        } else {
            ""
        };
        // Admit the whole entry (heading, root and replies) from an upper
        // bound of its escaped size before allocating any of it, so a single
        // oversized author or text cannot be expanded past the budget first.
        let mut bound = 2 * (sheet_name.len() + cell.len()) as u64 + 32;
        bound += quoted_comment_bound(
            comment["author"].as_str().unwrap_or("(unknown)"),
            comment["rootText"]
                .as_str()
                .or_else(|| comment["text"].as_str())
                .unwrap_or(""),
            ">",
        );
        if let Some(replies) = comment["replies"].as_array() {
            for reply in replies {
                bound = bound.saturating_add(quoted_comment_bound(
                    reply["author"].as_str().unwrap_or("(unknown)"),
                    reply["text"].as_str().unwrap_or(""),
                    ">>",
                ));
            }
        }
        let admitted = (out.len() as u64).saturating_add(bound);
        if admitted > limit {
            return Err(admitted);
        }
        let _ = writeln!(
            out,
            "### {} — {}{}\n",
            escape_text(sheet_name, TextContext::Heading),
            escape_text(cell, TextContext::Heading),
            status
        );
        let text = comment["rootText"]
            .as_str()
            .or_else(|| comment["text"].as_str())
            .unwrap_or("");
        write_quoted_comment(
            comment["author"].as_str().unwrap_or("(unknown)"),
            text,
            ">",
            "",
            out,
        );
        if let Some(replies) = comment["replies"].as_array() {
            for reply in replies {
                let status = if reply["resolved"].as_bool() == Some(true) {
                    " (resolved)"
                } else {
                    ""
                };
                write_quoted_comment(
                    reply["author"].as_str().unwrap_or("(unknown)"),
                    reply["text"].as_str().unwrap_or(""),
                    ">>",
                    status,
                    out,
                );
                if let Some(observed) = over(out) {
                    return Err(observed);
                }
            }
        }
        out.push('\n');
        if let Some(observed) = over(out) {
            return Err(observed);
        }
    }
    Ok(())
}

/// Upper bound of what [`write_quoted_comment`] appends: escaping at most
/// doubles the author and the text, and every text line gains the prefix, a
/// space and a newline.
fn quoted_comment_bound(author: &str, text: &str, prefix: &str) -> u64 {
    let lines = ooxml_common::markdown::line_count_bound(text);
    (2 * author.len() as u64)
        .saturating_add(2 * text.len() as u64)
        .saturating_add(lines.saturating_mul(prefix.len() as u64 + 2))
        .saturating_add(3 * prefix.len() as u64 + 32)
}

fn write_quoted_comment(author: &str, text: &str, prefix: &str, status: &str, out: &mut String) {
    let author = escape_text(&author.replace(['\r', '\n'], " "), TextContext::Inline);
    let _ = writeln!(out, "{prefix} **{author}**{status}");
    let _ = writeln!(out, "{prefix}");
    if text.is_empty() {
        let _ = writeln!(out, "{prefix}");
    } else {
        for line in split_lines(text) {
            let _ = writeln!(out, "{prefix} {}", escape_text(line, TextContext::Inline));
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    /// A `{"type":"shared","si":N}` cell must resolve to the sharedStrings
    /// table's text — the wire ships only the index, so markdown has to look
    /// it up (mirrors the runtime `t="s"` path).
    #[test]
    fn shared_cell_resolves_against_table() {
        let shared = vec![
            SharedString {
                text: "Alpha".to_string(),
                runs: None,
                ..Default::default()
            },
            SharedString {
                text: "Beta".to_string(),
                runs: None,
                ..Default::default()
            },
        ];
        let sheet = json!({
            "name": "Sheet1",
            "rows": [
                {
                    "index": 1,
                    "cells": [
                        { "col": 1, "row": 1, "value": { "type": "shared", "si": 1 } },
                        { "col": 2, "row": 1, "value": { "type": "number", "number": 3.0 } }
                    ]
                }
            ]
        });
        let mut out = String::new();
        render_sheet(&sheet, &shared, &mut out, u64::MAX).unwrap();
        assert!(
            out.contains("Beta"),
            "si=1 must resolve to shared[1] text, got:\n{out}"
        );
    }

    /// An out-of-range `si` resolves to empty text (historical fallback),
    /// leaving no populated cell — the sheet renders as empty.
    #[test]
    fn shared_cell_out_of_range_is_empty() {
        let shared: Vec<SharedString> = Vec::new();
        let sheet = json!({
            "name": "Sheet1",
            "rows": [
                {
                    "index": 1,
                    "cells": [
                        { "col": 1, "row": 1, "value": { "type": "shared", "si": 9 } }
                    ]
                }
            ]
        });
        let mut out = String::new();
        render_sheet(&sheet, &shared, &mut out, u64::MAX).unwrap();
        // No populated cells → only the heading is emitted (no table body).
        assert!(!out.contains('|'), "empty sheet must have no table: {out}");
    }

    #[test]
    fn review_comments_are_separated_quoted_and_threaded() {
        let sheet = json!({
            "name": "Review",
            "comments": [{
                "cellRef": "B4",
                "author": "Reviewer",
                "rootText": "Check this cell.\nKeep the formula.",
                "text": "Check this cell.\nKeep the formula.\nUpdated.",
                "resolved": true,
                "replies": [{ "author": "Editor", "text": "Updated." }]
            }]
        });
        let mut out = String::new();
        let mut has_comments = false;

        render_sheet_comments(&sheet, &mut out, &mut has_comments, u64::MAX).unwrap();

        assert!(out.starts_with("\n## Review comments\n"), "{out}");
        assert!(out.contains("### Review — B4 (resolved)"), "{out}");
        assert!(
            out.contains("> **Reviewer**\n>\n> Check this cell.\n> Keep the formula."),
            "{out}"
        );
        assert!(out.contains(">> **Editor**\n>>\n>> Updated."), "{out}");
    }

    /// Many comments on one sheet repeat its (possibly long) name in every
    /// heading; the budget is enforced while they are written.
    #[test]
    fn comments_stop_at_the_budget_while_rendering() {
        let comments: Vec<Value> = (1..=6000)
            .map(|row| json!({ "cellRef": format!("A{row}"), "author": "a", "text": "c" }))
            .collect();
        let sheet = json!({ "name": "n".repeat(16_384), "comments": comments });
        let mut out = String::new();
        let mut has_comments = false;
        let observed = render_sheet_comments(&sheet, &mut out, &mut has_comments, 1 << 20)
            .expect_err("over budget");
        assert!(observed > 1 << 20);
        assert!(out.len() < (1 << 20) + 64 * 1024, "{}", out.len());
    }

    /// The admission bound covers the worst-case expansion: every lone CR
    /// starts a prefixed line and every character is escaped.
    #[test]
    fn comment_bound_covers_lone_cr_lines_and_escaping() {
        for text in ["\r".repeat(512), "#\r".repeat(256), "&~".repeat(256)] {
            let mut out = String::new();
            write_quoted_comment("*_", &text, ">>", " (resolved)", &mut out);
            let bound = quoted_comment_bound("*_", &text, ">>");
            assert!(out.len() as u64 <= bound, "{} > {bound}", out.len());
        }
    }

    /// One oversized comment is rejected from its size bound before any of
    /// it is escaped or written.
    #[test]
    fn a_single_oversized_comment_is_admitted_before_writing() {
        let sheet = json!({
            "name": "S",
            "comments": [{ "cellRef": "A1", "author": "*".repeat(1 << 20), "text": "c" }],
        });
        let mut out = String::new();
        let mut has_comments = false;
        let observed = render_sheet_comments(&sheet, &mut out, &mut has_comments, 1 << 20)
            .expect_err("over budget");
        assert!(observed > 1 << 20);
        assert!(out.len() < 1024, "{}", out.len());
    }

    /// Many cells sharing one long shared string are admitted against the
    /// budget from borrowed display strings; nothing is copied per cell.
    #[test]
    fn repeated_shared_strings_are_admitted_before_expansion() {
        let shared = vec![SharedString {
            text: "s".repeat(32_767),
            runs: None,
            ..Default::default()
        }];
        let cells: Vec<Value> = (1..=8192)
            .map(|row| json!({ "col": 1, "row": row, "value": { "type": "shared", "si": 0 } }))
            .collect();
        let rows: Vec<Value> = cells
            .into_iter()
            .enumerate()
            .map(|(i, cell)| json!({ "index": i + 1, "cells": [cell] }))
            .collect();
        let sheet = json!({ "name": "S", "rows": rows });
        let mut out = String::new();
        let projected = render_sheet(&sheet, &shared, &mut out, 64 << 20).expect_err("over budget");
        assert!(projected > 64 << 20);
        assert!(out.len() < 1024, "{}", out.len());
    }

    /// A long sheet name on an otherwise empty sheet is charged before the
    /// empty-sheet early return.
    #[test]
    fn empty_sheet_headings_are_charged_to_the_budget() {
        let sheet = json!({ "name": "n".repeat(4096) });
        let mut out = String::new();
        assert_eq!(
            render_sheet(&sheet, &[], &mut out, 1024),
            Err(out.len() as u64)
        );
        let mut out = String::new();
        assert_eq!(render_sheet(&sheet, &[], &mut out, u64::MAX), Ok(()));
    }

    fn text_cell(row: u32, col: u32, text: &str) -> Value {
        json!({ "col": col, "row": row, "value": { "type": "text", "text": text } })
    }

    fn sheet_of(cells: &[(u32, u32, &str)], merges: Value) -> Value {
        let mut rows: BTreeMap<u32, Vec<Value>> = BTreeMap::new();
        for &(row, col, text) in cells {
            rows.entry(row).or_default().push(text_cell(row, col, text));
        }
        let rows: Vec<Value> = rows
            .into_iter()
            .map(|(index, cells)| json!({ "index": index, "cells": cells }))
            .collect();
        json!({ "name": "S", "rows": rows, "mergeCells": merges })
    }

    /// Continuation cells clear and top-left cells keep their value, including
    /// for duplicated and overlapping merges; gaps and empty middle rows keep
    /// their historical layout.
    #[test]
    fn merges_clear_continuation_cells_only() {
        let sheet = sheet_of(
            &[
                (1, 1, "a"),
                (1, 2, "b"),
                (2, 1, "c"),
                (2, 2, "d"),
                (4, 3, "e"),
                (6, 1, "f"),
            ],
            json!([
                { "top": 1, "left": 1, "bottom": 2, "right": 2 },
                { "top": 1, "left": 1, "bottom": 2, "right": 2 },
                { "top": 4, "left": 3, "bottom": 4, "right": 3 },
                { "top": 5, "left": 1, "bottom": 1048576, "right": 16384 }
            ]),
        );
        let mut out = String::new();
        render_sheet(&sheet, &[], &mut out, u64::MAX).unwrap();
        assert_eq!(
            out,
            "## S\n\n| a |  |  |\n| --- | --- | --- |\n|  |  | e |\n\n"
        );
    }

    /// Two far-apart cells span the whole grid without allocating it, and a
    /// sheet-wide merge is applied without enumerating its cells.
    #[test]
    fn far_apart_cells_and_sheet_wide_merges_stay_sparse() {
        let sheet = sheet_of(
            &[(1, 1, "a"), (1_048_576, 16_384, "z")],
            Value::Array(vec![
                json!({ "top": 1, "left": 1, "bottom": 1048576, "right": 16384 });
                1000
            ]),
        );
        let mut out = String::new();
        render_sheet(&sheet, &[], &mut out, u64::MAX).unwrap();
        assert_eq!(
            out.lines().count(),
            5,
            "heading, blank, header, separator, blank"
        );
        assert!(out.contains("| a |"));
        assert!(!out.contains('z'), "continuation cell must clear");

        let sheet = sheet_of(
            &(1..=64)
                .map(|row| (row, if row % 2 == 0 { 16_384 } else { 1 }, "x"))
                .collect::<Vec<_>>(),
            Value::Null,
        );
        let mut out = String::new();
        let projected = render_sheet(&sheet, &[], &mut out, 1 << 20).expect_err("over budget");
        assert!(projected > 1 << 20);

        // The projection is an upper bound of the written table: with the
        // limit at the projected size the table fits, one byte less fails.
        let sheet = sheet_of(&[(1, 1, "a"), (1, 16_384, "z")], Value::Null);
        let mut probe = String::new();
        render_sheet(&sheet, &[], &mut probe, u64::MAX).unwrap();
        let mut out = String::new();
        let heading = "## S\n\n".len() as u64;
        let projected = render_sheet(&sheet, &[], &mut out, heading).expect_err("over budget");
        assert!(
            projected >= probe.len() as u64,
            "{projected} < {}",
            probe.len()
        );
        let mut out = String::new();
        render_sheet(&sheet, &[], &mut out, projected).expect("fits its projection");
    }
}
