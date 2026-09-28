use crate::parsed_cache;
use rmcp::{handler::server::wrapper::Parameters, tool};
use schemars::JsonSchema;
use serde::Deserialize;
use serde_json::Value;
use xlsx_model::{Cell, CellRange, CellValue, MergeCell, Workbook, Worksheet};

// ─── Parameter types ─────────────────────────────────────────────────────────

#[derive(Debug, Deserialize, JsonSchema)]
pub struct XlsxPathParam {
    /// Absolute path to the XLSX file
    pub path: String,
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct XlsxSearchParam {
    /// Absolute path to the XLSX file
    pub path: String,
    /// Sheet name or 0-based index; omit to search all sheets
    pub sheet: Option<String>,
    /// Case-insensitive substring to search for in cell values and formulas
    pub query: String,
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct XlsxSheetParam {
    /// Absolute path to the XLSX file
    pub path: String,
    /// Sheet name (e.g. "Sheet1") or 0-based numeric index as a string (e.g. "0")
    pub sheet: String,
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct XlsxCellRangeParam {
    /// Absolute path to the XLSX file
    pub path: String,
    /// Sheet name or 0-based index
    pub sheet: String,
    /// Cell range in A1 notation, e.g. "A1:C10"
    pub range: String,
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct XlsxOptSheetParam {
    /// Absolute path to the XLSX file
    pub path: String,
    /// Sheet name or 0-based index; omit to scan all sheets
    pub sheet: Option<String>,
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct XlsxChartIndexParam {
    /// Absolute path to the XLSX file
    pub path: String,
    /// Sheet name or 0-based index
    pub sheet: String,
    /// 0-based chart index within the sheet (matches order in `xlsx_get_charts`)
    pub chart_index: usize,
}

// ─── Helpers ─────────────────────────────────────────────────────────────────

fn workbook(path: &str) -> Result<std::sync::Arc<parsed_cache::XlsxDocument>, String> {
    parsed_cache::xlsx(path)
}

fn resolve_sheet(workbook: &Workbook, identifier: &str) -> Result<(u32, String), String> {
    if let Ok(idx) = identifier.parse::<usize>() {
        let sheet = workbook.sheets.get(idx).ok_or_else(|| {
            format!(
                "sheet index {} out of range (total: {})",
                idx,
                workbook.sheets.len()
            )
        })?;
        return Ok((idx as u32, sheet.name.clone()));
    }
    workbook
        .sheets
        .iter()
        .enumerate()
        .find(|(_, sheet)| sheet.name == identifier)
        .map(|(idx, _)| (idx as u32, identifier.to_string()))
        .ok_or_else(|| {
            format!(
                "sheet '{}' not found (available: {})",
                identifier,
                workbook
                    .sheets
                    .iter()
                    .map(|sheet| sheet.name.as_str())
                    .collect::<Vec<_>>()
                    .join(", ")
            )
        })
}

fn target_sheets(
    workbook: &Workbook,
    identifier: Option<&str>,
) -> Result<Vec<(u32, String)>, String> {
    if let Some(identifier) = identifier {
        return resolve_sheet(workbook, identifier).map(|sheet| vec![sheet]);
    }
    Ok(workbook
        .sheets
        .iter()
        .enumerate()
        .map(|(idx, sheet)| (idx as u32, sheet.name.clone()))
        .collect())
}

enum LoadError {
    General(String),
    Sheet { name: String, source: String },
}

impl std::fmt::Display for LoadError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::General(error) => write!(f, "Error: {error}"),
            Self::Sheet { name, source } => write!(f, "Error parsing sheet '{name}': {source}"),
        }
    }
}

fn load_sheet(
    path: &str,
    identifier: &str,
) -> Result<(String, std::sync::Arc<Worksheet>), LoadError> {
    let doc = workbook(path).map_err(LoadError::General)?;
    let (idx, name) = resolve_sheet(&doc.workbook, identifier).map_err(LoadError::General)?;
    let sheet = parsed_cache::xlsx_sheet(path, &doc, idx, &name).map_err(LoadError::General)?;
    Ok((name, sheet))
}

fn load_targets(path: &str, identifier: Option<&str>) -> Result<TargetSheets, LoadError> {
    let doc = workbook(path).map_err(LoadError::General)?;
    let targets = target_sheets(&doc.workbook, identifier).map_err(LoadError::General)?;
    Ok(TargetSheets {
        path: path.into(),
        doc,
        targets: targets.into_iter(),
    })
}

struct TargetSheets {
    path: String,
    doc: std::sync::Arc<parsed_cache::XlsxDocument>,
    targets: std::vec::IntoIter<(u32, String)>,
}

impl Iterator for TargetSheets {
    type Item = Result<(String, std::sync::Arc<Worksheet>), LoadError>;

    fn next(&mut self) -> Option<Self::Item> {
        self.targets.next().map(|(idx, name)| {
            parsed_cache::xlsx_sheet(&self.path, &self.doc, idx, &name)
                .map(|sheet| (name.clone(), sheet))
                .map_err(|source| LoadError::Sheet { name, source })
        })
    }
}

fn parse_cell_ref(s: &str) -> Option<(u32, u32)> {
    let col_str: String = s.chars().take_while(|c| c.is_ascii_alphabetic()).collect();
    let row_str: String = s.chars().skip_while(|c| c.is_ascii_alphabetic()).collect();
    if col_str.is_empty() || row_str.is_empty() {
        return None;
    }
    let col = col_str
        .to_ascii_uppercase()
        .chars()
        .fold(0u32, |acc, c| acc * 26 + (c as u32 - 'A' as u32 + 1));
    Some((col, row_str.parse().ok()?))
}

fn col_to_letter(mut col: u32) -> String {
    let mut s = String::new();
    while col > 0 {
        col -= 1;
        s.insert(0, (b'A' + (col % 26) as u8) as char);
        col /= 26;
    }
    s
}

fn coords_to_a1(top: u32, left: u32, bottom: u32, right: u32) -> String {
    format!(
        "{}{}:{}{}",
        col_to_letter(left),
        top,
        col_to_letter(right),
        bottom
    )
}
fn merge_to_a1(merge: &MergeCell) -> String {
    coords_to_a1(merge.top, merge.left, merge.bottom, merge.right)
}
fn range_to_a1(range: &CellRange) -> String {
    coords_to_a1(range.top, range.left, range.bottom, range.right)
}

fn cell_display(cell: &Cell) -> String {
    match &cell.value {
        CellValue::Text { text, .. } => text.clone(),
        // The previous JSON projection turned non-finite numbers into null,
        // which displayed as empty and could not match a text search.
        CellValue::Number { number } if !number.is_finite() => String::new(),
        CellValue::Number { number } => {
            if number.fract() == 0.0 && number.abs() < 1e15 {
                format!("{}", *number as i64)
            } else {
                format!("{number}")
            }
        }
        CellValue::Bool { bool } => {
            if *bool {
                "TRUE".into()
            } else {
                "FALSE".into()
            }
        }
        CellValue::Error { error } => error.clone(),
        // The existing MCP display projection leaves shared-string references
        // unresolved. Keep its output stable independently of renderer behavior.
        CellValue::Empty | CellValue::Shared { .. } => String::new(),
    }
}

pub struct XlsxTools;

impl XlsxTools {
    #[tool(
        description = "Convert an XLSX file to GitHub-flavoured markdown — one `## SheetName` per sheet, followed by a pipe table of the cells' cached display values. Merged-cell continuation cells render empty. Formula cells show the cached result, not the formula text (use `xlsx_get_formulas` if you need the formulas). Designed for agents that need to *read* spreadsheet content efficiently. Lossy by design: drops styling, conditional formatting, charts, sparklines, drawings. For precise structure use the structured tools (`xlsx_get_cell_range`, `xlsx_get_sheet_layout`, etc.)"
    )]
    pub fn xlsx_to_markdown(Parameters(p): Parameters<XlsxPathParam>) -> String {
        match parsed_cache::markdown(
            &p.path,
            parsed_cache::MarkdownKind::Xlsx,
            xlsx_parser::to_markdown_native,
        ) {
            Ok(md) => md,
            Err(e) => format!("Error: {e}"),
        }
    }

    #[tool(
        description = "Parse an XLSX file and return workbook overview including sheet names and IDs"
    )]
    pub fn xlsx_parse(Parameters(p): Parameters<XlsxPathParam>) -> String {
        match workbook(&p.path) {
            Ok(doc) => {
                serde_json::to_string(&doc.workbook).unwrap_or_else(|e| format!("Error: {e}"))
            }
            Err(e) => format!("Error: {e}"),
        }
    }

    #[tool(description = "Return the dimensions (max row and column) of a worksheet")]
    pub fn xlsx_get_sheet_dimensions(Parameters(p): Parameters<XlsxSheetParam>) -> String {
        let (name, ws) = match load_sheet(&p.path, &p.sheet) {
            Ok(x) => x,
            Err(e) => return e.to_string(),
        };
        let max_row = ws.rows.iter().map(|row| row.index).max().unwrap_or(0);
        let max_col = ws
            .rows
            .iter()
            .flat_map(|row| &row.cells)
            .map(|cell| cell.col)
            .max()
            .unwrap_or(0);
        serde_json::json!({ "sheet": name, "maxRow": max_row, "maxCol": max_col, "maxColLetter": col_to_letter(max_col) }).to_string()
    }

    #[tool(
        description = "Return cell values and formulas for a given range (e.g. \"A1:C10\") in a worksheet"
    )]
    pub fn xlsx_get_cell_range(Parameters(p): Parameters<XlsxCellRangeParam>) -> String {
        let (name, ws) = match load_sheet(&p.path, &p.sheet) {
            Ok(x) => x,
            Err(e) => return e.to_string(),
        };
        let parts: Vec<&str> = p.range.split(':').collect();
        if parts.len() != 2 {
            return format!("Error: range must be in 'A1:C10' format, got '{}'", p.range);
        }
        let (c1, r1) = match parse_cell_ref(parts[0]) {
            Some(v) => v,
            None => return format!("Error: invalid cell reference '{}'", parts[0]),
        };
        let (c2, r2) = match parse_cell_ref(parts[1]) {
            Some(v) => v,
            None => return format!("Error: invalid cell reference '{}'", parts[1]),
        };
        let (row_min, row_max) = (r1.min(r2), r1.max(r2));
        let (col_min, col_max) = (c1.min(c2), c1.max(c2));
        let rows: Vec<Value> = ws.rows.iter().filter(|row| (row_min..=row_max).contains(&row.index)).map(|row| {
            let cells: Vec<Value> = row.cells.iter().filter(|cell| (col_min..=col_max).contains(&cell.col)).map(|cell| {
                let mut entry = serde_json::json!({ "ref": format!("{}{}", col_to_letter(cell.col), row.index), "value": cell_display(cell) });
                if let Some(formula) = &cell.formula { entry["formula"] = Value::String(formula.clone()); }
                entry
            }).collect();
            serde_json::json!({ "row": row.index, "cells": cells })
        }).collect();
        serde_json::json!({ "sheet": name, "range": p.range, "rows": rows }).to_string()
    }

    #[tool(description = "Return all cells that contain formulas in a worksheet")]
    pub fn xlsx_get_formulas(Parameters(p): Parameters<XlsxSheetParam>) -> String {
        let (name, ws) = match load_sheet(&p.path, &p.sheet) {
            Ok(x) => x,
            Err(e) => return e.to_string(),
        };
        let formulas: Vec<Value> = ws.rows.iter().flat_map(|row| row.cells.iter().filter_map(move |cell| cell.formula.as_ref().map(|formula| {
            serde_json::json!({ "ref": format!("{}{}", col_to_letter(cell.col), row.index), "formula": formula, "cachedValue": cell_display(cell) })
        }))).collect();
        serde_json::json!({ "sheet": name, "formulas": formulas }).to_string()
    }

    #[tool(
        description = "Search for a substring in cell values and formulas across one or all sheets of an XLSX file"
    )]
    pub fn xlsx_search_cells(Parameters(p): Parameters<XlsxSearchParam>) -> String {
        let sheets = match load_targets(&p.path, p.sheet.as_deref()) {
            Ok(x) => x,
            Err(e) => return e.to_string(),
        };
        let query_lower = p.query.to_lowercase();
        let mut matches = Vec::new();
        for sheet in sheets {
            let (name, ws) = match sheet {
                Ok(sheet) => sheet,
                Err(e) => return e.to_string(),
            };
            for row in &ws.rows {
                for cell in &row.cells {
                    let value = cell_display(cell);
                    let formula = cell.formula.as_deref().unwrap_or("");
                    if value.to_lowercase().contains(&query_lower)
                        || formula.to_lowercase().contains(&query_lower)
                    {
                        let mut entry = serde_json::json!({ "sheet": name, "ref": format!("{}{}", col_to_letter(cell.col), row.index), "value": value });
                        if !formula.is_empty() {
                            entry["formula"] = Value::String(formula.into());
                        }
                        matches.push(entry);
                    }
                }
            }
        }
        serde_json::json!({ "query": p.query, "matchCount": matches.len(), "matches": matches })
            .to_string()
    }

    #[tool(
        description = "List charts on a worksheet (or all sheets if `sheet` is omitted). Returns a summary per chart: anchor cell range, chart type, title, axes, legend, and a series outline (without numeric values)"
    )]
    pub fn xlsx_get_charts(Parameters(p): Parameters<XlsxOptSheetParam>) -> String {
        let sheets = match load_targets(&p.path, p.sheet.as_deref()) {
            Ok(x) => x,
            Err(e) => return e.to_string(),
        };
        let mut all_charts = Vec::new();
        for sheet in sheets {
            let (name, ws) = match sheet {
                Ok(sheet) => sheet,
                Err(e) => return e.to_string(),
            };
            for (chart_idx, anchor) in ws.charts.iter().enumerate() {
                let chart = &anchor.chart;
                let series: Vec<Value> = chart
                    .series
                    .iter()
                    .map(|s| {
                        serde_json::json!({
                            "name": s.name, "type": s.series_type, "color": s.color,
                            "showMarker": s.show_marker, "valueCount": s.values.len(),
                        })
                    })
                    .collect();
                all_charts.push(serde_json::json!({
                "sheet": name, "chartIndex": chart_idx,
                "anchor": { "from": { "col": anchor.from_col, "row": anchor.from_row }, "to": { "col": anchor.to_col, "row": anchor.to_row } },
                "type": chart.chart_type, "barDir": null, "grouping": null,
                "title": chart.title, "legend": { "show": chart.show_legend, "position": chart.legend_pos },
                "axes": { "cat": { "title": chart.cat_axis_title, "formatCode": chart.cat_axis_format_code, "hidden": chart.cat_axis_hidden },
                           "val": { "title": chart.val_axis_title, "formatCode": chart.val_axis_format_code, "hidden": chart.val_axis_hidden } },
                "categories": chart.categories, "seriesCount": series.len(), "series": series,
            }));
            }
        }
        serde_json::json!({ "charts": all_charts }).to_string()
    }

    #[tool(
        description = "Return one chart's full series data (categories and per-point values) for drill-down. `chart_index` matches the index from `xlsx_get_charts` for the same sheet"
    )]
    pub fn xlsx_get_chart_series(Parameters(p): Parameters<XlsxChartIndexParam>) -> String {
        let (name, ws) = match load_sheet(&p.path, &p.sheet) {
            Ok(x) => x,
            Err(e) => return e.to_string(),
        };
        let Some(anchor) = ws.charts.get(p.chart_index) else {
            return format!(
                "Error: chart index {} out of range (total: {})",
                p.chart_index,
                ws.charts.len()
            );
        };
        let chart = &anchor.chart;
        let series: Vec<Value> = chart
            .series
            .iter()
            .map(|s| {
                serde_json::json!({
                    "name": s.name, "type": s.series_type, "color": s.color, "values": s.values,
                    "categories": s.categories, "valFormatCode": s.val_format_code,
                })
            })
            .collect();
        serde_json::json!({ "sheet": name, "chartIndex": p.chart_index, "type": chart.chart_type,
            "title": chart.title, "categories": chart.categories, "series": series })
        .to_string()
    }

    #[tool(
        description = "Return all defined names (named ranges) visible in the workbook. Includes workbook-global names plus each sheet's local names; duplicates across sheets are merged"
    )]
    pub fn xlsx_get_named_ranges(Parameters(p): Parameters<XlsxPathParam>) -> String {
        let sheets = match load_targets(&p.path, None) {
            Ok(x) => x,
            Err(e) => return e.to_string(),
        };
        let mut seen: Vec<(String, String, String)> = Vec::new();
        for sheet in sheets {
            let (name, ws) = match sheet {
                Ok(sheet) => sheet,
                Err(e) => return e.to_string(),
            };
            for dn in &ws.defined_names {
                if !seen
                    .iter()
                    .any(|(n, f, _)| n == &dn.name && f == &dn.formula)
                {
                    seen.push((dn.name.clone(), dn.formula.clone(), name.clone()));
                }
            }
        }
        let defined_names: Vec<Value> = seen
            .into_iter()
            .map(|(name, formula, sheet)| {
                serde_json::json!({
                    "name": name, "refersTo": formula, "firstSeenSheet": sheet,
                })
            })
            .collect();
        serde_json::json!({ "definedNames": defined_names }).to_string()
    }

    #[tool(
        description = "List Excel Tables (Ctrl+T tables, ECMA-376 §18.5) on a sheet or across all sheets. Returns each table's range, style, header/totals row counts"
    )]
    pub fn xlsx_get_tables(Parameters(p): Parameters<XlsxOptSheetParam>) -> String {
        let sheets = match load_targets(&p.path, p.sheet.as_deref()) {
            Ok(x) => x,
            Err(e) => return e.to_string(),
        };
        let mut tables = Vec::new();
        for sheet in sheets {
            let (name, ws) = match sheet {
                Ok(sheet) => sheet,
                Err(e) => return e.to_string(),
            };
            for table in &ws.tables {
                tables.push(serde_json::json!({
                    "sheet": name, "range": range_to_a1(&table.range), "styleName": table.style_name,
                    "headerRowCount": table.header_row_count, "totalsRowCount": table.totals_row_count,
                    "showRowStripes": table.show_row_stripes, "showColumnStripes": table.show_column_stripes,
                }));
            }
        }
        serde_json::json!({ "tables": tables }).to_string()
    }

    #[tool(
        description = "Return all merged cell ranges on a worksheet as A1 strings (e.g. \"A1:B2\")"
    )]
    pub fn xlsx_get_merged_cells(Parameters(p): Parameters<XlsxSheetParam>) -> String {
        let (name, ws) = match load_sheet(&p.path, &p.sheet) {
            Ok(x) => x,
            Err(e) => return e.to_string(),
        };
        let merges: Vec<String> = ws.merge_cells.iter().map(merge_to_a1).collect();
        serde_json::json!({ "sheet": name, "merges": merges }).to_string()
    }

    #[tool(
        description = "Return conditional formatting rules on a worksheet. Each entry has the affected ranges (sqref) and the rule body (CellIs, Expression, ColorScale, DataBar, Top10, AboveAverage, IconSet, Other)"
    )]
    pub fn xlsx_get_conditional_formats(Parameters(p): Parameters<XlsxSheetParam>) -> String {
        let (name, ws) = match load_sheet(&p.path, &p.sheet) {
            Ok(x) => x,
            Err(e) => return e.to_string(),
        };
        let formats: Vec<Value> = ws.conditional_formats.iter().map(|cf| serde_json::json!({
            "ranges": cf.sqref.iter().map(range_to_a1).collect::<Vec<_>>(), "rules": cf.rules,
        })).collect();
        serde_json::json!({ "sheet": name, "formats": formats }).to_string()
    }

    #[tool(
        description = "Return all `<dataValidation>` rules on a worksheet: affected ranges, type, operator, formulas, and the optional prompt / error messages"
    )]
    pub fn xlsx_get_data_validations(Parameters(p): Parameters<XlsxSheetParam>) -> String {
        let (name, ws) = match load_sheet(&p.path, &p.sheet) {
            Ok(x) => x,
            Err(e) => return e.to_string(),
        };
        serde_json::json!({ "sheet": name, "validations": ws.data_validations }).to_string()
    }

    #[tool(
        description = "Return all comments on a worksheet (or all sheets if `sheet` is omitted) with full text and resolved author. Each entry: { sheet, cellRef, author?, text }"
    )]
    pub fn xlsx_get_comments(Parameters(p): Parameters<XlsxOptSheetParam>) -> String {
        let sheets = match load_targets(&p.path, p.sheet.as_deref()) {
            Ok(x) => x,
            Err(e) => return e.to_string(),
        };
        let mut comments = Vec::new();
        for sheet in sheets {
            let (name, ws) = match sheet {
                Ok(sheet) => sheet,
                Err(e) => return e.to_string(),
            };
            for comment in &ws.comments {
                comments.push(serde_json::json!({
                    "sheet": name, "cellRef": comment.cell_ref, "author": comment.author, "text": comment.text,
                }));
            }
        }
        serde_json::json!({ "comments": comments }).to_string()
    }

    #[tool(
        description = "Return per-sheet layout: explicit column widths, row heights, freeze panes, gridline visibility, default sizes, and tab color"
    )]
    pub fn xlsx_get_sheet_layout(Parameters(p): Parameters<XlsxSheetParam>) -> String {
        let (name, ws) = match load_sheet(&p.path, &p.sheet) {
            Ok(x) => x,
            Err(e) => return e.to_string(),
        };
        let cols: Vec<Value> = ws
            .col_widths
            .iter()
            .map(|(col, width)| {
                serde_json::json!({
                    "col": col, "width": width, "letter": col_to_letter(*col),
                })
            })
            .collect();
        let rows: Vec<Value> = ws
            .row_heights
            .iter()
            .map(|(row, height)| {
                serde_json::json!({
                    "row": row, "height": height,
                })
            })
            .collect();
        serde_json::json!({
            "sheet": name, "defaultColWidth": ws.default_col_width, "defaultRowHeight": ws.default_row_height,
            "freeze": { "rows": ws.freeze_rows, "cols": ws.freeze_cols },
            "showGridlines": ws.show_gridlines, "showZeros": ws.show_zeros, "tabColor": ws.tab_color,
            "colWidths": cols, "rowHeights": rows,
        }).to_string()
    }
}

#[cfg(test)]
mod sample_tests {
    use super::*;
    use std::io::{Cursor, Read, Write};
    use zip::{write::SimpleFileOptions, ZipArchive, ZipWriter};

    #[test]
    fn numeric_cells_keep_blank_non_finite_values_and_exact_finite_display() {
        let source = include_bytes!(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../xlsx/public/demo/sample-1.xlsx"
        ));
        let mut archive = ZipArchive::new(Cursor::new(source.as_slice())).unwrap();
        let mut output = ZipWriter::new(Cursor::new(Vec::new()));
        for index in 0..archive.len() {
            let mut part = archive.by_index(index).unwrap();
            let name = part.name().to_string();
            let mut bytes = Vec::new();
            if name == "xl/worksheets/sheet1.xml" {
                bytes.extend_from_slice(
                    br#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1">
    <c r="A1"><v>NaN</v></c><c r="B1"><f>1/0</f><v>inf</v></c>
    <c r="C1"><v>-inf</v></c><c r="D1"><v>123456789012345.67</v></c>
  </row></sheetData>
</worksheet>"#,
                );
            } else {
                part.read_to_end(&mut bytes).unwrap();
            }
            output
                .start_file(name, SimpleFileOptions::default())
                .unwrap();
            output.write_all(&bytes).unwrap();
        }
        let path = std::env::temp_dir().join(format!(
            "ooxml-mcp-numeric-{}-{}.xlsx",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::SystemTime::UNIX_EPOCH)
                .unwrap()
                .as_nanos()
        ));
        std::fs::write(&path, output.finish().unwrap().into_inner()).unwrap();
        let path = path.to_str().unwrap();

        let range: Value = serde_json::from_str(&XlsxTools::xlsx_get_cell_range(Parameters(
            XlsxCellRangeParam {
                path: path.into(),
                sheet: "0".into(),
                range: "A1:D1".into(),
            },
        )))
        .unwrap();
        let cells = range["rows"][0]["cells"].as_array().unwrap();
        assert_eq!(cells.len(), 4);
        assert_eq!(cells[0]["value"], "");
        assert_eq!(cells[1]["value"], "");
        assert_eq!(cells[2]["value"], "");
        assert_eq!(cells[3]["value"], "123456789012345.67");

        let formulas: Value =
            serde_json::from_str(&XlsxTools::xlsx_get_formulas(Parameters(XlsxSheetParam {
                path: path.into(),
                sheet: "0".into(),
            })))
            .unwrap();
        assert_eq!(formulas["formulas"][0]["ref"], "B1");
        assert_eq!(formulas["formulas"][0]["cachedValue"], "");

        for query in ["nan", "inf"] {
            let found: Value =
                serde_json::from_str(&XlsxTools::xlsx_search_cells(Parameters(XlsxSearchParam {
                    path: path.into(),
                    sheet: Some("0".into()),
                    query: query.into(),
                })))
                .unwrap();
            assert_eq!(found["matchCount"], 0, "unexpected match for {query}");
        }
        let found: Value =
            serde_json::from_str(&XlsxTools::xlsx_search_cells(Parameters(XlsxSearchParam {
                path: path.into(),
                sheet: Some("0".into()),
                query: "123456789012345.67".into(),
            })))
            .unwrap();
        assert_eq!(found["matches"][0]["ref"], "D1");
        assert_eq!(found["matches"][0]["value"], "123456789012345.67");
        std::fs::remove_file(path).unwrap();
    }

    #[test]
    fn sheet_errors_keep_single_and_all_sheet_wording() {
        let source = include_bytes!(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../xlsx/public/demo/sample-1.xlsx"
        ));
        let mut archive = ZipArchive::new(Cursor::new(source.as_slice())).unwrap();
        let mut output = ZipWriter::new(Cursor::new(Vec::new()));
        for index in 0..archive.len() {
            let mut part = archive.by_index(index).unwrap();
            if part.name() == "xl/_rels/workbook.xml.rels" {
                continue;
            }
            let name = part.name().to_string();
            let mut bytes = Vec::new();
            part.read_to_end(&mut bytes).unwrap();
            output
                .start_file(name, SimpleFileOptions::default())
                .unwrap();
            output.write_all(&bytes).unwrap();
        }
        let path = std::env::temp_dir().join(format!(
            "ooxml-mcp-missing-rels-{}-{}.xlsx",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::SystemTime::UNIX_EPOCH)
                .unwrap()
                .as_nanos()
        ));
        std::fs::write(&path, output.finish().unwrap().into_inner()).unwrap();
        let path = path.to_str().unwrap();
        let single = XlsxTools::xlsx_get_cell_range(Parameters(XlsxCellRangeParam {
            path: path.into(),
            sheet: "0".into(),
            range: "A1:C10".into(),
        }));
        assert_eq!(single, "Error: entry not found: xl/_rels/workbook.xml.rels");
        let all = XlsxTools::xlsx_get_charts(Parameters(XlsxOptSheetParam {
            path: path.into(),
            sheet: None,
        }));
        assert_eq!(
            all,
            "Error parsing sheet 'Dashboard': entry not found: xl/_rels/workbook.xml.rels"
        );
        std::fs::remove_file(path).unwrap();
    }

    fn sample_path() -> String {
        format!(
            "{}/../xlsx/public/demo/sample-1.xlsx",
            env!("CARGO_MANIFEST_DIR")
        )
    }

    fn pp(path: &str) -> Parameters<XlsxPathParam> {
        Parameters(XlsxPathParam { path: path.into() })
    }

    #[test]
    fn xlsx_to_markdown_sample() {
        let path = sample_path();
        if !std::path::Path::new(&path).exists() {
            return;
        }
        let out = XlsxTools::xlsx_to_markdown(pp(&path));
        assert!(!out.starts_with("Error:"), "errored: {out}");
        assert!(
            out.contains("## "),
            "missing sheet heading: {}",
            &out[..200.min(out.len())]
        );
        assert!(out.contains("|"), "no pipe table emitted");
    }

    #[test]
    fn xlsx_parse_sample_returns_workbook() {
        let path = sample_path();
        if !std::path::Path::new(&path).exists() {
            return; // sample missing in this checkout — skip.
        }
        let out = XlsxTools::xlsx_parse(pp(&path));
        let v: Value = serde_json::from_str(&out).expect("xlsx_parse must return JSON");
        assert!(
            v["sheets"].as_array().is_some(),
            "missing 'sheets' array: {out}"
        );
    }

    #[test]
    fn xlsx_get_charts_sample_returns_charts_field() {
        let path = sample_path();
        if !std::path::Path::new(&path).exists() {
            return;
        }
        let out = XlsxTools::xlsx_get_charts(Parameters(XlsxOptSheetParam {
            path: path.clone(),
            sheet: None,
        }));
        let v: Value = serde_json::from_str(&out).expect("must return JSON");
        assert!(
            v["charts"].as_array().is_some(),
            "missing 'charts' array: {out}"
        );
    }

    #[test]
    fn xlsx_get_named_ranges_sample() {
        let path = sample_path();
        if !std::path::Path::new(&path).exists() {
            return;
        }
        let out = XlsxTools::xlsx_get_named_ranges(pp(&path));
        let v: Value = serde_json::from_str(&out).expect("must return JSON");
        assert!(
            v["definedNames"].as_array().is_some(),
            "missing 'definedNames'"
        );
    }

    #[test]
    fn xlsx_get_merged_cells_first_sheet() {
        let path = sample_path();
        if !std::path::Path::new(&path).exists() {
            return;
        }
        let out = XlsxTools::xlsx_get_merged_cells(Parameters(XlsxSheetParam {
            path: path.clone(),
            sheet: "0".into(),
        }));
        let v: Value = serde_json::from_str(&out).expect("must return JSON");
        assert!(v["merges"].as_array().is_some(), "missing 'merges'");
    }

    #[test]
    fn xlsx_get_sheet_layout_first_sheet() {
        let path = sample_path();
        if !std::path::Path::new(&path).exists() {
            return;
        }
        let out = XlsxTools::xlsx_get_sheet_layout(Parameters(XlsxSheetParam {
            path: path.clone(),
            sheet: "0".into(),
        }));
        let v: Value = serde_json::from_str(&out).expect("must return JSON");
        assert!(v["sheet"].is_string(), "missing 'sheet' name in {out}");
        assert!(v["colWidths"].as_array().is_some(), "missing 'colWidths'");
        assert!(v["rowHeights"].as_array().is_some(), "missing 'rowHeights'");
    }

    #[test]
    fn xlsx_get_data_validations_smoke() {
        let path = sample_path();
        if !std::path::Path::new(&path).exists() {
            return;
        }
        let out = XlsxTools::xlsx_get_data_validations(Parameters(XlsxSheetParam {
            path,
            sheet: "0".into(),
        }));
        let v: Value = serde_json::from_str(&out).expect("must return JSON");
        assert!(
            v["validations"].as_array().is_some(),
            "missing 'validations'"
        );
    }

    #[test]
    fn xlsx_get_comments_smoke() {
        let path = sample_path();
        if !std::path::Path::new(&path).exists() {
            return;
        }
        let out = XlsxTools::xlsx_get_comments(Parameters(XlsxOptSheetParam { path, sheet: None }));
        let v: Value = serde_json::from_str(&out).expect("must return JSON");
        assert!(v["comments"].as_array().is_some(), "missing 'comments'");
    }

    #[test]
    fn xlsx_invalid_path_returns_error_string() {
        let out = XlsxTools::xlsx_parse(pp("/nonexistent/does-not-exist.xlsx"));
        assert!(out.starts_with("Error:"), "expected error, got: {out}");
    }
}
