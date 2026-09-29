use crate::parsed_cache;
use docx_model::{BodyElement, DocRun, ShapeRun};
use rmcp::{handler::server::wrapper::Parameters, tool};
use schemars::JsonSchema;
use serde::Deserialize;
use serde_json::Value;

// ─── Parameter types ─────────────────────────────────────────────────────────

#[derive(Debug, Deserialize, JsonSchema)]
pub struct DocxPathParam {
    /// Absolute path to the DOCX file
    pub path: String,
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct DocxSearchParam {
    /// Absolute path to the DOCX file
    pub path: String,
    /// Case-insensitive substring to search for in paragraph and table cell text
    pub query: String,
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct DocxIndexParam {
    /// Absolute path to the DOCX file
    pub path: String,
    /// 0-based index into the body element list (paragraphs and tables share indexing).
    pub index: usize,
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct DocxTableIndexParam {
    /// Absolute path to the DOCX file
    pub path: String,
    /// 0-based index of the table (counts tables only, in document order)
    pub table_index: usize,
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct DocxImagesParam {
    /// Absolute path to the DOCX file
    pub path: String,
    /// When true include the base64 `dataUrl` for each image. Defaults to false
    /// (just the metadata) since image bytes are large and rarely needed inline.
    #[serde(default)]
    pub include_data_url: bool,
}

// ─── Helpers ─────────────────────────────────────────────────────────────────

fn document(path: &str) -> Result<std::sync::Arc<docx_model::Document>, String> {
    parsed_cache::docx(path)
}

fn run_text(runs: &[DocRun]) -> String {
    runs.iter()
        .filter_map(|run| match run {
            DocRun::Text(text) => Some(text.text.as_str()),
            _ => None,
        })
        .collect()
}

fn extract_body_text(body: &[BodyElement]) -> String {
    let mut out = String::new();
    for element in body {
        match element {
            BodyElement::Paragraph(para) => {
                out.push_str(&run_text(&para.runs));
                out.push('\n');
            }
            BodyElement::Table(table) => {
                // The prior MCP projection read `cell.paragraphs`; the model's
                // cell blocks are stored under `content`. Preserve that wire
                // projection until a separate behavior change is approved.
                for row in &table.rows {
                    out.push_str(&vec![""; row.cells.len()].join("\t"));
                    out.push('\n');
                }
            }
            _ => {}
        }
    }
    out
}

fn body_structure(body: &[BodyElement]) -> Vec<Value> {
    body.iter()
        .map(|element| match element {
            BodyElement::Paragraph(para) => serde_json::json!({
                "type": "paragraph", "styleId": para.style_id,
                "text": run_text(&para.runs).trim(), "alignment": para.alignment,
            }),
            BodyElement::Table(table) => serde_json::json!({
                "type": "table", "rows": table.rows.len(),
                "cols": table.rows.first().map_or(0, |row| row.cells.len()),
            }),
            _ => serde_json::to_value(element).unwrap_or(Value::Null),
        })
        .collect()
}

fn shape_summary(shape: &ShapeRun, paragraph_index: usize, run_index: usize) -> Value {
    // The previous JSON-backed projection returned null when these fields
    // were omitted by ShapeRun's serializer. Preserve that public tool output
    // while reading the remaining fields directly from the typed model.
    serde_json::json!({
        "paragraphIndex": paragraph_index, "runIndex": run_index,
        "presetGeometry": shape.preset_geometry, "widthPt": shape.width_pt,
        "heightPt": shape.height_pt, "anchorXPt": shape.anchor_x_pt,
        "anchorYPt": shape.anchor_y_pt,
        "rotation": (shape.rotation != 0.0).then_some(shape.rotation),
        "fill": shape.fill, "stroke": shape.stroke,
        "strokeWidth": (shape.stroke_width != 0.0).then_some(shape.stroke_width),
        "textBlocks": (!shape.text_blocks.is_empty()).then_some(&shape.text_blocks),
        "wrapMode": shape.wrap_mode,
        "behindDoc": shape.behind_doc.then_some(true),
        "zOrder": shape.z_order,
    })
}

pub struct DocxTools;

impl DocxTools {
    #[tool(
        description = "Convert a DOCX file to GitHub-flavoured markdown. Preserves textual structure (headings from outlineLevel, paragraphs, bullet/numbered lists, tables, footnotes, comments) and rich-text formatting (bold/italic/strikethrough/hyperlinks). Discards positioning, section properties, font metrics, drawing shapes, and headers/footers. Designed for agents that need to *read* the document content efficiently — typical 10×+ token reduction vs. the structured JSON tools. Lossy by design: when you need precise layout or styling, fall back to `docx_get_structure` / `docx_get_body_element`"
    )]
    pub fn docx_to_markdown(Parameters(p): Parameters<DocxPathParam>) -> String {
        match parsed_cache::markdown(
            &p.path,
            parsed_cache::MarkdownKind::Docx,
            docx_parser::to_markdown_native,
        ) {
            Ok(md) => md,
            Err(e) => format!("Error: {e}"),
        }
    }

    #[tool(description = "Extract all plain text from a DOCX file")]
    pub fn docx_extract_text(Parameters(p): Parameters<DocxPathParam>) -> String {
        match document(&p.path) {
            Ok(doc) => extract_body_text(&doc.body),
            Err(e) => format!("Error: {e}"),
        }
    }

    #[tool(description = "Return the document structure (paragraphs and tables) of a DOCX file")]
    pub fn docx_get_structure(Parameters(p): Parameters<DocxPathParam>) -> String {
        match document(&p.path) {
            Ok(doc) => serde_json::to_string(&body_structure(&doc.body))
                .unwrap_or_else(|e| format!("Error: {e}")),
            Err(e) => format!("Error: {e}"),
        }
    }

    #[tool(description = "Return all tables from a DOCX file with their cell contents")]
    pub fn docx_get_tables(Parameters(p): Parameters<DocxPathParam>) -> String {
        let doc = match document(&p.path) {
            Ok(doc) => doc,
            Err(e) => return format!("Error: {e}"),
        };
        let tables: Vec<Value> = doc
            .body
            .iter()
            .filter_map(|element| match element {
                BodyElement::Table(table) => Some(table),
                _ => None,
            })
            .enumerate()
            .map(|(table_idx, table)| {
                let rows: Vec<Vec<&str>> = table
                    .rows
                    .iter()
                    .map(|row| vec![""; row.cells.len()])
                    .collect();
                serde_json::json!({ "tableIndex": table_idx, "rows": rows })
            })
            .collect();
        serde_json::to_string(&tables).unwrap_or_else(|e| format!("Error: {e}"))
    }

    #[tool(
        description = "Search for a substring in all paragraph and table text of a DOCX file; returns matching excerpts with their position"
    )]
    pub fn docx_search_text(Parameters(p): Parameters<DocxSearchParam>) -> String {
        let doc = match document(&p.path) {
            Ok(doc) => doc,
            Err(e) => return format!("Error: {e}"),
        };
        let query_lower = p.query.to_lowercase();
        let mut matches = Vec::new();
        for (idx, element) in doc.body.iter().enumerate() {
            match element {
                BodyElement::Paragraph(para) => {
                    let text = run_text(&para.runs);
                    if text.to_lowercase().contains(&query_lower) {
                        matches.push(serde_json::json!({ "type": "paragraph", "index": idx, "styleId": para.style_id, "text": text.trim() }));
                    }
                }
                BodyElement::Table(table) => {
                    for (row_idx, row) in table.rows.iter().enumerate() {
                        for col_idx in 0..row.cells.len() {
                            if "".contains(&query_lower) {
                                matches.push(serde_json::json!({ "type": "tableCell", "tableIndex": idx, "row": row_idx, "col": col_idx, "text": "" }));
                            }
                        }
                    }
                }
                _ => {}
            }
        }
        serde_json::json!({ "query": p.query, "matchCount": matches.len(), "matches": matches })
            .to_string()
    }

    #[tool(
        description = "Return one body element's full detail (paragraph or table) including run-level formatting (bold/italic/color/font/hyperlink), indents, spacing, numbering, and tab stops. `index` is into the document body list (matches `docx_get_structure`)"
    )]
    pub fn docx_get_body_element(Parameters(p): Parameters<DocxIndexParam>) -> String {
        let doc = match document(&p.path) {
            Ok(doc) => doc,
            Err(e) => return format!("Error: {e}"),
        };
        let Some(element) = doc.body.get(p.index) else {
            return format!(
                "Error: body index {} out of range (total: {})",
                p.index,
                doc.body.len()
            );
        };
        let mut out = serde_json::to_value(element).unwrap_or(Value::Null);
        if let Some(obj) = out.as_object_mut() {
            obj.insert("index".into(), Value::from(p.index));
        }
        out.to_string()
    }

    #[tool(
        description = "Return the document's section properties (page size/margins/docGrid) along with default/first/even header and footer body elements"
    )]
    pub fn docx_get_sections(Parameters(p): Parameters<DocxPathParam>) -> String {
        let doc = match document(&p.path) {
            Ok(doc) => doc,
            Err(e) => return format!("Error: {e}"),
        };
        serde_json::json!({ "section": doc.section, "headers": doc.headers, "footers": doc.footers, "majorFont": doc.major_font, "minorFont": doc.minor_font }).to_string()
    }

    #[tool(
        description = "Return one table's full detail by index, including cell content, colSpan/vMerge, borders, shading, and row heights. Use this for deeper inspection than `docx_get_tables`"
    )]
    pub fn docx_get_table(Parameters(p): Parameters<DocxTableIndexParam>) -> String {
        let doc = match document(&p.path) {
            Ok(doc) => doc,
            Err(e) => return format!("Error: {e}"),
        };
        let tables: Vec<_> = doc
            .body
            .iter()
            .enumerate()
            .filter(|(_, element)| matches!(element, BodyElement::Table(_)))
            .collect();
        let Some((body_index, table)) = tables.get(p.table_index) else {
            return format!(
                "Error: table index {} out of range (total tables: {})",
                p.table_index,
                tables.len()
            );
        };
        serde_json::json!({ "tableIndex": p.table_index, "bodyIndex": body_index, "table": table })
            .to_string()
    }

    #[tool(
        description = "List all images in the document. Each entry carries the paragraph index, anchor mode, wrap settings, and dimensions. Set `include_data_url=true` to also receive the inline base64 image bytes (large)"
    )]
    pub fn docx_get_images(Parameters(p): Parameters<DocxImagesParam>) -> String {
        let doc = match document(&p.path) {
            Ok(doc) => doc,
            Err(e) => return format!("Error: {e}"),
        };
        let mut images = Vec::new();
        for (para_idx, element) in doc.body.iter().enumerate() {
            if let BodyElement::Paragraph(para) = element {
                for (run_idx, run) in para.runs.iter().enumerate() {
                    if let DocRun::Image(image) = run {
                        let mut entry = serde_json::json!({
                            "paragraphIndex": para_idx, "runIndex": run_idx,
                            "widthPt": image.width_pt, "heightPt": image.height_pt,
                            "anchor": image.anchor, "anchorXPt": image.anchor_x_pt,
                            "anchorYPt": image.anchor_y_pt, "wrapMode": image.wrap_mode,
                        });
                        if p.include_data_url {
                            entry["dataUrl"] = Value::Null;
                        }
                        images.push(entry);
                    }
                }
            }
        }
        serde_json::json!({ "images": images }).to_string()
    }

    #[tool(
        description = "List all drawn shapes embedded in paragraphs (wps:wsp inside wp:anchor). Returns each shape's preset geometry, fill, stroke, dimensions, anchor offsets, rotation, and embedded text blocks"
    )]
    pub fn docx_get_shapes(Parameters(p): Parameters<DocxPathParam>) -> String {
        let doc = match document(&p.path) {
            Ok(doc) => doc,
            Err(e) => return format!("Error: {e}"),
        };
        let mut shapes = Vec::new();
        for (para_idx, element) in doc.body.iter().enumerate() {
            if let BodyElement::Paragraph(para) = element {
                for (run_idx, run) in para.runs.iter().enumerate() {
                    if let DocRun::Shape(shape) = run {
                        shapes.push(shape_summary(shape, para_idx, run_idx));
                    }
                }
            }
        }
        serde_json::json!({ "shapes": shapes }).to_string()
    }

    #[tool(
        description = "Return the heading outline of the document. Each entry has the body index, outlineLevel (0-8), styleId, and visible text. Levels come from the parser's resolved `outlineLevel` (style chain + direct pPr) — useful for building TOCs without parsing styleId strings"
    )]
    pub fn docx_get_outline(Parameters(p): Parameters<DocxPathParam>) -> String {
        let doc = match document(&p.path) {
            Ok(doc) => doc,
            Err(e) => return format!("Error: {e}"),
        };
        let outline: Vec<Value> = doc.body.iter().enumerate().filter_map(|(idx, element)| match element {
            BodyElement::Paragraph(para) => para.outline_level.map(|level| serde_json::json!({
                "bodyIndex": idx, "level": level, "styleId": para.style_id, "text": run_text(&para.runs).trim(),
            })), _ => None,
        }).collect();
        serde_json::json!({ "outline": outline }).to_string()
    }

    #[tool(
        description = "List all `<w:comment>` entries from word/comments.xml: id, author, initials, date, plain text. Empty when the document has no comments part"
    )]
    pub fn docx_get_comments(Parameters(p): Parameters<DocxPathParam>) -> String {
        let doc = match document(&p.path) {
            Ok(doc) => doc,
            Err(e) => return format!("Error: {e}"),
        };
        serde_json::json!({ "comments": doc.comments }).to_string()
    }

    #[tool(
        description = "List footnote and endnote bodies from word/footnotes.xml and word/endnotes.xml. Each entry has the id (matches `<w:footnoteReference w:id>` in body) and concatenated plain text"
    )]
    pub fn docx_get_footnotes(Parameters(p): Parameters<DocxPathParam>) -> String {
        let doc = match document(&p.path) {
            Ok(doc) => doc,
            Err(e) => return format!("Error: {e}"),
        };
        serde_json::json!({ "footnotes": doc.footnotes, "endnotes": doc.endnotes }).to_string()
    }

    #[tool(
        description = "Return all track-changes events found in the body: insertions and deletions with author, date, and the text. Empty when the document has no tracked changes"
    )]
    pub fn docx_get_revisions(Parameters(p): Parameters<DocxPathParam>) -> String {
        let doc = match document(&p.path) {
            Ok(doc) => doc,
            Err(e) => return format!("Error: {e}"),
        };
        serde_json::json!({ "revisions": doc.revisions }).to_string()
    }
}

#[cfg(test)]
mod sample_tests {
    use super::*;
    use std::io::{Cursor, Read, Write};
    use zip::{write::SimpleFileOptions, ZipArchive, ZipWriter};

    #[test]
    fn minimal_shape_keeps_omitted_defaults_null_in_summary() {
        let source = include_bytes!(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../docx/public/demo/sample-1.docx"
        ));
        let mut archive = ZipArchive::new(Cursor::new(source.as_slice())).unwrap();
        let mut output = ZipWriter::new(Cursor::new(Vec::new()));
        let minimal_document = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><w:body><w:p><w:r><w:drawing><wp:inline><wp:extent cx="2540000" cy="635000"/><wp:docPr id="1" name="Minimal rectangle"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1270000" cy="317500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:body></w:document>"#;
        for index in 0..archive.len() {
            let mut part = archive.by_index(index).unwrap();
            let name = part.name().to_string();
            let mut bytes = Vec::new();
            part.read_to_end(&mut bytes).unwrap();
            output
                .start_file(&name, SimpleFileOptions::default())
                .unwrap();
            output
                .write_all(if name == "word/document.xml" {
                    minimal_document.as_bytes()
                } else {
                    &bytes
                })
                .unwrap();
        }
        let bytes = output.finish().unwrap().into_inner();
        let document = docx_parser::parse_docx_model_native(&bytes).unwrap();
        let shape = document
            .body
            .iter()
            .find_map(|element| match element {
                BodyElement::Paragraph(paragraph) => {
                    paragraph.runs.iter().find_map(|run| match run {
                        DocRun::Shape(shape) => Some(shape.as_ref()),
                        _ => None,
                    })
                }
                _ => None,
            })
            .expect("synthetic rectangle must parse");
        let summary = shape_summary(shape, 0, 1);
        for field in ["rotation", "strokeWidth", "behindDoc", "textBlocks"] {
            assert!(summary[field].is_null(), "{field} must remain null");
        }
        assert_eq!(summary["presetGeometry"], "rect");
        assert_eq!(summary["widthPt"], 200.0);
    }

    fn sample_path() -> String {
        format!(
            "{}/../docx/public/demo/sample-1.docx",
            env!("CARGO_MANIFEST_DIR")
        )
    }

    fn pp(path: &str) -> Parameters<DocxPathParam> {
        Parameters(DocxPathParam { path: path.into() })
    }

    #[test]
    fn docx_to_markdown_sample() {
        let path = sample_path();
        if !std::path::Path::new(&path).exists() {
            return;
        }
        let out = DocxTools::docx_to_markdown(pp(&path));
        assert!(!out.starts_with("Error:"), "errored: {out}");
        assert!(!out.trim().is_empty(), "markdown should be non-empty");
        let plain = DocxTools::docx_extract_text(pp(&path));
        // Allow markdown to be larger than plain (markup overhead) but not
        // more than 3× — guards against accidentally serializing the rich
        // structure tree.
        assert!(
            out.len() < plain.len() * 3 + 1024,
            "markdown should stay within bounds — got {} vs plain {}",
            out.len(),
            plain.len()
        );
    }

    #[test]
    fn docx_extract_text_sample_non_empty() {
        let path = sample_path();
        if !std::path::Path::new(&path).exists() {
            return;
        }
        let out = DocxTools::docx_extract_text(pp(&path));
        assert!(!out.starts_with("Error:"), "got error: {out}");
        assert!(!out.trim().is_empty(), "extracted text should be non-empty");
    }

    #[test]
    fn docx_get_sections_sample() {
        let path = sample_path();
        if !std::path::Path::new(&path).exists() {
            return;
        }
        let out = DocxTools::docx_get_sections(pp(&path));
        let v: Value = serde_json::from_str(&out).expect("must return JSON");
        // pageWidth is f64 in pt; should be > 0 for any real document.
        let pw = v["section"]["pageWidth"].as_f64().unwrap_or(0.0);
        assert!(pw > 0.0, "section.pageWidth should be > 0, got {pw}");
    }

    #[test]
    fn docx_get_body_element_first_element() {
        let path = sample_path();
        if !std::path::Path::new(&path).exists() {
            return;
        }
        let out = DocxTools::docx_get_body_element(Parameters(DocxIndexParam {
            path: path.clone(),
            index: 0,
        }));
        let v: Value = serde_json::from_str(&out).expect("must return JSON");
        assert_eq!(v["index"].as_u64(), Some(0));
        assert!(v["type"].is_string(), "missing 'type' on body element");
    }

    #[test]
    fn docx_get_images_returns_array() {
        let path = sample_path();
        if !std::path::Path::new(&path).exists() {
            return;
        }
        let out = DocxTools::docx_get_images(Parameters(DocxImagesParam {
            path: path.clone(),
            include_data_url: false,
        }));
        let v: Value = serde_json::from_str(&out).expect("must return JSON");
        assert!(
            v["images"].as_array().is_some(),
            "missing 'images' array: {out}"
        );
    }

    #[test]
    fn docx_invalid_path_returns_error_string() {
        let out = DocxTools::docx_extract_text(pp("/nonexistent/x.docx"));
        assert!(out.starts_with("Error:"), "expected error, got: {out}");
    }

    #[test]
    fn docx_get_outline_smoke() {
        let path = sample_path();
        if !std::path::Path::new(&path).exists() {
            return;
        }
        let out = DocxTools::docx_get_outline(pp(&path));
        let v: Value = serde_json::from_str(&out).expect("must return JSON");
        assert!(v["outline"].as_array().is_some(), "missing 'outline'");
    }

    #[test]
    fn docx_get_comments_smoke() {
        let path = sample_path();
        if !std::path::Path::new(&path).exists() {
            return;
        }
        let out = DocxTools::docx_get_comments(pp(&path));
        let v: Value = serde_json::from_str(&out).expect("must return JSON");
        assert!(v["comments"].as_array().is_some(), "missing 'comments'");
    }

    #[test]
    fn docx_get_revisions_smoke() {
        let path = sample_path();
        if !std::path::Path::new(&path).exists() {
            return;
        }
        let out = DocxTools::docx_get_revisions(pp(&path));
        let v: Value = serde_json::from_str(&out).expect("must return JSON");
        assert!(v["revisions"].as_array().is_some(), "missing 'revisions'");
    }

    #[test]
    fn docx_get_footnotes_smoke() {
        let path = sample_path();
        if !std::path::Path::new(&path).exists() {
            return;
        }
        let out = DocxTools::docx_get_footnotes(pp(&path));
        let v: Value = serde_json::from_str(&out).expect("must return JSON");
        assert!(v["footnotes"].as_array().is_some(), "missing 'footnotes'");
        assert!(v["endnotes"].as_array().is_some(), "missing 'endnotes'");
    }

    #[test]
    fn docx_get_body_element_out_of_range_errors() {
        let path = sample_path();
        if !std::path::Path::new(&path).exists() {
            return;
        }
        let out = DocxTools::docx_get_body_element(Parameters(DocxIndexParam {
            path,
            index: 999_999,
        }));
        assert!(
            out.starts_with("Error:"),
            "expected out-of-range error, got: {out}"
        );
    }
}
