//! RED production-boundary tests for the finite, caller-adjustable worksheet
//! resource policy. The fixture is a tiny synthetic OPC package generated in
//! this module. It has 2 rows, 3 cells, an inline string and a formula. No
//! sibling-module private helpers, mocks or source-text checks are used.

use std::io::{Cursor, Write};

use zip::write::SimpleFileOptions;
use zip::ZipWriter;

use super::{
    parse_sheet_native, parse_sheet_native_with_worksheet_limits, write_test_content_types,
    WorksheetResourcePolicy,
};

const MAX_SAFE: u64 = 9_007_199_254_740_991;
const BIG: u64 = MAX_SAFE - 1;

const SHEET: &str = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1" t="inlineStr"><is><t>ab</t></is></c></row><row r="2"><c r="A2"><f>A1*2</f><v>2</v></c></row></sheetData></worksheet>"#;

fn fixture() -> Vec<u8> {
    let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
    write_test_content_types(&mut writer);
    let options = SimpleFileOptions::default();
    let parts: [(&str, &str); 4] = [
        (
            "_rels/.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#,
        ),
        (
            "xl/workbook.xml",
            r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
        ),
        (
            "xl/_rels/workbook.xml.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
        ),
        ("xl/worksheets/sheet1.xml", SHEET),
    ];
    for (path, body) in parts {
        writer.start_file(path, options).expect("in-memory ZIP");
        writer.write_all(body.as_bytes()).expect("in-memory ZIP");
    }
    writer.finish().expect("in-memory ZIP").into_inner()
}

fn policy(rows: u64, cells: u64, owned: u64, json: u64) -> WorksheetResourcePolicy {
    WorksheetResourcePolicy::new(rows, cells, owned, json).expect("valid finite policy")
}

fn parse(policy: WorksheetResourcePolicy) -> Result<String, String> {
    parse_sheet_native_with_worksheet_limits(&fixture(), 0, "Sheet1", policy)
}

fn legacy() -> String {
    parse_sheet_native(&fixture(), 0, "Sheet1").expect("legacy default parse succeeds")
}

fn assert_adjustable(error: &str, needles: &[&str]) {
    assert!(
        error.contains(r#""configurable":true"#),
        "adjustable policy error must be configurable: {error}"
    );
    assert!(
        !error.contains(r#""configurable":false"#),
        "adjustable policy error must not claim a hard guard: {error}"
    );
    for needle in needles {
        assert!(error.contains(needle), "missing {needle} in {error}");
    }
}

#[test]
fn rows_and_cells_limits_are_adjustable_and_policy_values_are_finite() {
    let baseline = legacy();
    assert_eq!(
        parse(WorksheetResourcePolicy::default()).expect("default policy succeeds"),
        baseline
    );
    assert_eq!(
        parse(policy(100_000, 250_000, 33_554_432, 67_108_864)).expect("explicit defaults"),
        baseline
    );

    let rows_error = parse(policy(1, BIG, BIG, BIG)).expect_err("maxRows 1 rejects 2 rows");
    assert_adjustable(&rows_error, &[r#""limit":1"#, r#""observed":2"#]);
    assert!(rows_error.to_lowercase().contains("row"), "{rows_error}");

    let cells_error = parse(policy(BIG, 2, BIG, BIG)).expect_err("maxCells 2 rejects 3 cells");
    assert_adjustable(&cells_error, &[r#""limit":2"#, r#""observed":3"#]);
    assert!(cells_error.to_lowercase().contains("cell"), "{cells_error}");

    assert_eq!(parse(policy(2, 3, BIG, BIG)).expect("exact boundary"), baseline);
    assert_eq!(parse(policy(3, 4, BIG, BIG)).expect("increased limits"), baseline);

    for invalid in [0, MAX_SAFE] {
        assert!(WorksheetResourcePolicy::new(invalid, 1, 1, 1).is_err());
        assert!(WorksheetResourcePolicy::new(1, invalid, 1, 1).is_err());
        assert!(WorksheetResourcePolicy::new(1, 1, invalid, 1).is_err());
        assert!(WorksheetResourcePolicy::new(1, 1, 1, invalid).is_err());
    }
    // Largest safe policy is accepted and does not pre-allocate to its limits.
    assert_eq!(parse(policy(BIG, BIG, BIG, BIG)).expect("max safe minus one"), baseline);
}

#[test]
fn owned_utf8_limit_counts_materialized_strings_and_formula_at_exact_boundary() {
    let baseline = legacy();
    // Two "number" discriminators, one "text" discriminator, inline "ab",
    // and the independently owned formula "A1*2". Object keys are not charged.
    let expected = 6 + 6 + 4 + 2 + 4;

    assert_eq!(
        parse(policy(BIG, BIG, expected, BIG)).expect("exact owned boundary"),
        baseline
    );
    let error = parse(policy(BIG, BIG, expected - 1, BIG)).expect_err("one under rejects");
    let limit = format!(r#""limit":{}"#, expected - 1);
    assert_adjustable(&error, &[limit.as_str()]);
    assert!(error.to_lowercase().contains("utf8") || error.to_lowercase().contains("utf-8"), "{error}");
}

#[test]
fn worksheet_json_limit_measures_exact_native_serialization_only() {
    let baseline = legacy();
    let exact = baseline.len() as u64;
    assert_eq!(
        parse(policy(BIG, BIG, BIG, exact)).expect("exact JSON boundary"),
        baseline
    );
    let error = parse(policy(BIG, BIG, BIG, exact - 1)).expect_err("one under rejects");
    let limit = format!(r#""limit":{}"#, exact - 1);
    assert_adjustable(&error, &[limit.as_str(), "worksheet-json"]);
    let payload: serde_json::Value = serde_json::from_str(
        error.strip_prefix("OOXML_RESOURCE_LIMIT:").expect("typed resource envelope")
    ).expect("resource error JSON");
    assert_eq!(payload["details"]["stage"], "serialization");
    assert_eq!(payload["details"]["violation"]["resource"], "worksheet-json");
}
