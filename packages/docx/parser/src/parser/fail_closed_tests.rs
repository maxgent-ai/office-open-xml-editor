use super::*;
use std::io::{Cursor, Write};

fn package(body: &str, parts: &[(&str, &str)]) -> Vec<u8> {
    let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
    write_test_content_types(&mut writer);
    let document = format!(
        r#"<w:document xmlns:w="{W_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>{body}</w:body></w:document>"#
    );
    for (name, xml) in
        std::iter::once(("word/document.xml", document.as_str())).chain(parts.iter().copied())
    {
        writer
            .start_file(name, zip::write::SimpleFileOptions::default())
            .unwrap();
        writer.write_all(xml.as_bytes()).unwrap();
    }
    writer.finish().unwrap().into_inner()
}

fn assert_depth_error(error: &str, limit: u32) {
    let payload: serde_json::Value =
        serde_json::from_str(error.strip_prefix("OOXML_RESOURCE_LIMIT:").expect(error)).unwrap();
    assert_eq!(payload["code"], "ooxml-resource-limit");
    assert_eq!(payload["details"]["violation"]["metric"], "depth");
    assert_eq!(payload["details"]["violation"]["limit"], limit);
}

#[test]
fn optional_styles_and_header_footer_depth_fail_closed_at_package_boundary() {
    let deep = "<w:x>".repeat(257) + &"</w:x>".repeat(257);
    for (tag, path) in [
        ("styles", "word/styles.xml"),
        ("hdr", "word/header1.xml"),
        ("ftr", "word/footer1.xml"),
    ] {
        let xml = format!(r#"<w:{tag} xmlns:w="{W_NS}">{deep}</w:{tag}>"#);
        let kind = if tag == "hdr" { "header" } else { "footer" };
        let rels = format!(
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="h" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/{kind}" Target="{}"/></Relationships>"#,
            path.strip_prefix("word/").unwrap()
        );
        let body =
            format!(r#"<w:p/><w:sectPr><w:{kind}Reference w:type="default" r:id="h"/></w:sectPr>"#);
        let bytes = package(
            &body,
            &[(path, &xml), ("word/_rels/document.xml.rels", &rels)],
        );
        let error = parse_from_bytes(&bytes).expect_err("optional part limit must reject document");
        assert_depth_error(&error, 256);
        let error = parse_from_bytes_streamed_with_limits(&bytes, None, None, "depth-test")
            .expect_err("streamed document must reject optional part limit");
        assert_depth_error(&error, 256);
    }
}

#[test]
fn nested_table_depth_fails_closed_before_publishing_document() {
    std::thread::Builder::new()
        .stack_size(32 * 1024 * 1024)
        .spawn(|| {
            let mut body = "<w:p/>".to_string();
            for _ in 0..66 {
                body = format!("<w:tbl><w:tr><w:tc>{body}</w:tc></w:tr></w:tbl>");
            }
            let bytes = package(&body, &[]);
            let error = parse_from_bytes(&bytes).expect_err("nested tables must reject document");
            assert_depth_error(&error, 64);
            let error = parse_from_bytes_streamed_with_limits(&bytes, None, None, "depth-test")
                .expect_err("streamed tables must reject document");
            assert_depth_error(&error, 64);
        })
        .unwrap()
        .join()
        .unwrap();
}

#[test]
fn textbox_container_and_table_depth_fail_closed() {
    let xml = format!(
        r#"<w:txbxContent xmlns:w="{W_NS}"><w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl></w:txbxContent>"#
    );
    let document = roxmltree::Document::parse(&xml).unwrap();
    for container_descent in [true, false] {
        let mut zip = Zip::new(Cursor::new(package("<w:p/>", &[]))).unwrap();
        let error = zip
            .run_operation("textbox-test", |_| {
                let parse = if container_descent {
                    parse_text_box_content_at_depth
                } else {
                    parse_text_box_content
                };
                let _provisional = parse(
                    document.root_element(),
                    &StyleMap::parse(""),
                    &mut NumberingMap::default(),
                    &HashMap::new(),
                    &ChartMap::default(),
                    &HashMap::new(),
                    &ThemeColors::default(),
                    DepthGuard::with_limit(0),
                );
                Ok(())
            })
            .unwrap_err();
        assert_depth_error(&error, 0);
    }
}

mod lexical_depth_cases {
    include!("../../../../ooxml-common/tests/support/lexical_depth_cases.rs");
}

#[test]
fn optional_styles_distinguish_lexical_depth_from_shallow_content_errors() {
    let baseline = package("<w:p/>", &[("word/styles.xml", "<broken")]);
    let expected = serde_json::to_value(parse_from_bytes(&baseline).unwrap()).unwrap();
    let expected_streamed = serde_json::to_value(
        parse_from_bytes_streamed_with_limits(&baseline, None, None, "depth-test").unwrap(),
    )
    .unwrap();
    for (name, xml, crossing) in lexical_depth_cases::lexical_depth_cases() {
        let bytes = package("<w:p/>", &[("word/styles.xml", &xml)]);
        let result = parse_from_bytes(&bytes);
        let streamed = parse_from_bytes_streamed_with_limits(&bytes, None, None, "depth-test");
        if crossing {
            assert_depth_error(&result.unwrap_err(), 256);
            assert_depth_error(&streamed.unwrap_err(), 256);
        } else {
            assert_eq!(
                serde_json::to_value(result.expect(name)).unwrap(),
                expected,
                "{name}"
            );
            assert_eq!(
                serde_json::to_value(streamed.expect(name)).unwrap(),
                expected_streamed,
                "{name}"
            );
        }
    }
}

fn assert_hard_limit(error: &str, resource: &str, limit: u64) {
    let payload: serde_json::Value =
        serde_json::from_str(error.strip_prefix("OOXML_RESOURCE_LIMIT:").expect(error)).unwrap();
    let violation = &payload["details"]["violation"];
    assert_eq!(violation["resource"], resource, "{error}");
    assert_eq!(violation["metric"], "depth");
    assert_eq!(violation["limit"], limit);
    assert_eq!(violation["observed"], limit + 1);
    assert_eq!(violation["configurable"], false);
}

/// `basedOn` chains of `levels` distinct paragraph styles; the body uses the
/// most-derived one so every paragraph resolves the whole chain.
fn chained_styles_package(levels: usize) -> Vec<u8> {
    let mut styles = format!(r#"<w:styles xmlns:w="{W_NS}">"#);
    for level in 0..levels {
        let based_on = if level == 0 {
            String::new()
        } else {
            format!(r#"<w:basedOn w:val="S{}"/>"#, level - 1)
        };
        styles.push_str(&format!(
            r#"<w:style w:type="paragraph" w:styleId="S{level}">{based_on}</w:style>"#
        ));
    }
    styles.push_str("</w:styles>");
    let body = format!(
        r#"<w:p><w:pPr><w:pStyle w:val="S{}"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>"#,
        levels - 1
    );
    package(&body, &[("word/styles.xml", &styles)])
}

#[test]
fn style_inheritance_depth_fails_closed_above_the_hop_limit() {
    let limit = crate::styles::MAX_STYLE_INHERITANCE_DEPTH;
    parse_from_bytes(&chained_styles_package(limit)).expect("a chain at the limit parses");
    let bytes = chained_styles_package(limit + 1);
    let error = parse_from_bytes(&bytes).expect_err("an over-long chain must reject the document");
    assert_hard_limit(&error, "style-inheritance", limit as u64);
    let error = parse_from_bytes_streamed_with_limits(&bytes, None, None, "style-depth-test")
        .expect_err("streamed document must reject the over-long chain");
    assert_hard_limit(&error, "style-inheritance", limit as u64);
}

/// `count` unclosed complex fields; `begin` fldChars are flat run siblings,
/// so XML depth limits never see them.
fn open_fields_package(count: usize) -> Vec<u8> {
    let runs = r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>REF a</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>"#
        .repeat(count);
    package(&format!("<w:p>{runs}<w:r><w:t>x</w:t></w:r></w:p>"), &[])
}

#[test]
fn open_complex_field_count_fails_closed_above_the_limit() {
    parse_from_bytes(&open_fields_package(MAX_OPEN_COMPLEX_FIELDS))
        .expect("fields at the limit parse");
    let bytes = open_fields_package(MAX_OPEN_COMPLEX_FIELDS + 1);
    let error = parse_from_bytes(&bytes).expect_err("too many open fields must reject");
    assert_hard_limit(&error, "complex-field", MAX_OPEN_COMPLEX_FIELDS as u64);
    let error = parse_from_bytes_streamed_with_limits(&bytes, None, None, "field-depth-test")
        .expect_err("streamed document must reject too many open fields");
    assert_hard_limit(&error, "complex-field", MAX_OPEN_COMPLEX_FIELDS as u64);
}

#[test]
fn incremental_toc_scan_matches_whole_instruction_classification() {
    // Oracle: the previous whole-text predicate, re-run after every chunk.
    fn classify_toc(instruction: &str) -> bool {
        instruction
            .split_whitespace()
            .next()
            .is_some_and(|token| token.eq_ignore_ascii_case("TOC"))
    }
    let cases: &[&[&str]] = &[
        &[" TOC \\o \"1-3\""],
        &["T", "O", "C", " \\h"],
        &["  ", "\u{3000}toc", "X"],
        &["TO", "CX"],
        &["TOCX"],
        &["", " ", "PAGEREF _Toc1 \\h"],
        &["ＴOC"],
        &["\ttoc\n"],
    ];
    for chunks in cases {
        let mut scan = FirstTokenScan::default();
        let mut text = String::new();
        for chunk in *chunks {
            text.push_str(chunk);
            assert_eq!(
                scan.feed_is_toc(chunk),
                classify_toc(&text),
                "{chunks:?} at {text:?}"
            );
        }
    }
}

#[test]
fn table_style_inheritance_is_bounded_even_when_ids_are_reused() {
    // A later character style reusing each table style id replaces only the
    // paragraph/character map; the table map still holds the long chain.
    let limit = crate::styles::MAX_STYLE_INHERITANCE_DEPTH;
    let mut styles = format!(r#"<w:styles xmlns:w="{W_NS}">"#);
    for level in 0..=limit {
        let based_on = if level == 0 {
            String::new()
        } else {
            format!(r#"<w:basedOn w:val="T{}"/>"#, level - 1)
        };
        styles.push_str(&format!(
            r#"<w:style w:type="table" w:styleId="T{level}">{based_on}</w:style><w:style w:type="character" w:styleId="T{level}"/>"#
        ));
    }
    styles.push_str("</w:styles>");
    let body = format!(
        r#"<w:tbl><w:tblPr><w:tblStyle w:val="T{limit}"/></w:tblPr><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl><w:p/>"#
    );
    let bytes = package(&body, &[("word/styles.xml", &styles)]);
    let error = parse_from_bytes(&bytes).expect_err("long table chain must reject");
    assert_hard_limit(&error, "style-inheritance", limit as u64);
}

#[test]
fn many_begin_markers_in_one_run_fail_closed_without_rescanning_the_run() {
    let begins = r#"<w:fldChar w:fldCharType="begin"/>"#.repeat(MAX_OPEN_COMPLEX_FIELDS + 1);
    let bytes = package(&format!("<w:p><w:r>{begins}</w:r></w:p>"), &[]);
    let error = parse_from_bytes(&bytes).expect_err("too many open fields must reject");
    assert_hard_limit(&error, "complex-field", MAX_OPEN_COMPLEX_FIELDS as u64);
}

fn checkbox_marks(document: &Document) -> String {
    serde_json::to_string(document)
        .unwrap()
        .chars()
        .filter(|c| matches!(c, '☐' | '☒'))
        .collect()
}

#[test]
fn checkbox_facts_stay_scoped_to_their_own_run_in_every_parse_path() {
    let field = |checked: bool| {
        format!(
            r#"<w:p><w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:checkBox><w:default w:val="{}"/></w:checkBox></w:ffData></w:fldChar></w:r><w:r><w:instrText>FORMCHECKBOX</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>"#,
            u8::from(checked)
        )
    };
    let body: String = [false, true, false, true, false, true]
        .into_iter()
        .map(field)
        .collect();
    let bytes = package(&body, &[]);
    let expected = "☐☒☐☒☐☒";
    assert_eq!(checkbox_marks(&parse_from_bytes(&bytes).unwrap()), expected);
    let streamed = parse_from_bytes_streamed_with_limits(&bytes, None, None, "checkbox").unwrap();
    assert_eq!(checkbox_marks(&streamed), expected);
}

#[test]
fn repeated_separators_and_balanced_fields_in_one_run_stay_linear() {
    // One field with many `separate` markers: only the first divides the
    // instruction from the result; the stored result text still renders.
    let separators = r#"<w:fldChar w:fldCharType="separate"/>"#.repeat(4_000);
    let body = format!(
        r#"<w:p><w:r><w:fldChar w:fldCharType="begin"/><w:instrText>{} </w:instrText>{separators}<w:t>result</w:t><w:fldChar w:fldCharType="end"/></w:r></w:p>"#,
        "REF a ".repeat(4_000)
    );
    let document = parse_from_bytes(&package(&body, &[])).expect("repeated separators parse");
    assert!(serde_json::to_string(&document).unwrap().contains("result"));

    // Many balanced fields with visible text segments in a single run.
    let sequence = r#"<w:fldChar w:fldCharType="begin"/><w:instrText>REF a</w:instrText><w:fldChar w:fldCharType="separate"/><w:t>x</w:t><w:fldChar w:fldCharType="end"/>"#;
    let body = format!("<w:p><w:r>{}</w:r></w:p>", sequence.repeat(4_000));
    let started = std::time::Instant::now();
    let document = parse_from_bytes(&package(&body, &[])).expect("balanced fields parse");
    let text = serde_json::to_string(&document).unwrap();
    assert!(text.matches("\"x\"").count() >= 1);
    // Linear work: well under the quadratic-era seconds even in debug.
    assert!(started.elapsed() < std::time::Duration::from_secs(20));
}

#[test]
fn large_run_properties_are_resolved_once_for_many_visible_segments() {
    let rpr = format!("<w:rPr>{}<w:b/></w:rPr>", "<w:noProof/>".repeat(4_000));
    let sequence = r#"<w:fldChar w:fldCharType="begin"/><w:instrText>REF a</w:instrText><w:fldChar w:fldCharType="separate"/><w:t>x</w:t><w:fldChar w:fldCharType="end"/>"#;
    let body = format!("<w:p><w:r>{rpr}{}</w:r></w:p>", sequence.repeat(4_000));
    let started = std::time::Instant::now();
    let document = parse_from_bytes(&package(&body, &[])).expect("fields with large rPr parse");
    assert!(serde_json::to_string(&document).unwrap().contains("\"x\""));
    assert!(started.elapsed() < std::time::Duration::from_secs(20));
}

#[test]
fn style_depth_validation_stops_at_the_limit_despite_many_other_styles() {
    // A very deep chain beside many independent styles: the validator must
    // reject without re-walking the deep chain's state for every style.
    let mut styles = format!(r#"<w:styles xmlns:w="{W_NS}">"#);
    for level in 0..20_000 {
        let based_on = if level == 0 {
            String::new()
        } else {
            format!(r#"<w:basedOn w:val="D{}"/>"#, level - 1)
        };
        styles.push_str(&format!(
            r#"<w:style w:type="paragraph" w:styleId="D{level}">{based_on}</w:style><w:style w:type="paragraph" w:styleId="I{level}"/>"#
        ));
    }
    styles.push_str("</w:styles>");
    let bytes = package("<w:p/>", &[("word/styles.xml", &styles)]);
    let started = std::time::Instant::now();
    let error = parse_from_bytes(&bytes).expect_err("deep chain rejected");
    let limit = crate::styles::MAX_STYLE_INHERITANCE_DEPTH;
    assert_hard_limit(&error, "style-inheritance", limit as u64);
    assert!(started.elapsed() < std::time::Duration::from_secs(20));
}

#[test]
fn hidden_field_segments_share_one_formatting_copy_and_name_the_styles_part() {
    // A hidden run with a large retained property and many visible segments.
    let font = "F".repeat(1 << 20);
    let sequence =
        r#"<w:t>x</w:t><w:fldChar w:fldCharType="begin"/><w:fldChar w:fldCharType="end"/>"#;
    let body = format!(
        r#"<w:p><w:r><w:rPr><w:rFonts w:ascii="{font}"/><w:vanish/></w:rPr>{}</w:r></w:p>"#,
        sequence.repeat(4_000)
    );
    let started = std::time::Instant::now();
    parse_from_bytes(&package(&body, &[])).expect("hidden fields parse");
    assert!(started.elapsed() < std::time::Duration::from_secs(20));

    // The depth error names the relationship's styles part.
    let limit = crate::styles::MAX_STYLE_INHERITANCE_DEPTH;
    let mut styles = format!(r#"<w:styles xmlns:w="{W_NS}">"#);
    for level in 0..=limit {
        let based_on = if level == 0 {
            String::new()
        } else {
            format!(r#"<w:basedOn w:val="S{}"/>"#, level - 1)
        };
        styles.push_str(&format!(
            r#"<w:style w:type="paragraph" w:styleId="S{level}">{based_on}</w:style>"#
        ));
    }
    styles.push_str("</w:styles>");
    let rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles2.xml"/></Relationships>"#;
    let bytes = package(
        "<w:p/>",
        &[
            ("word/styles2.xml", &styles),
            ("word/_rels/document.xml.rels", rels),
        ],
    );
    let error = parse_from_bytes(&bytes).expect_err("over-long chain rejected");
    let payload: serde_json::Value =
        serde_json::from_str(error.strip_prefix("OOXML_RESOURCE_LIMIT:").expect(&error)).unwrap();
    assert_eq!(payload["details"]["violation"]["part"], "word/styles2.xml");
}
