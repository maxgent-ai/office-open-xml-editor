//! Public synthetic DOC bytes exercise the real style/PAPX/piece cascade and
//! canonical projection; these are not Office fidelity fixtures.
use super::table_tests::with_papx;
use super::tests::{
    default_note_dop, passive_special_source, source_with_stories, source_with_typography,
    NotesFixture,
};
use crate::cfb::{test_support::build_scoped_cfb, CompoundFile};
use docx_model::{BodyElement, Document};

fn base(text: &str) -> Vec<u8> {
    source_with_typography(
        text,
        &[(text.encode_utf16().count(), 2, 12240, 15840, 1, 720)],
        None,
        None,
        None,
        None,
    )
}

fn with_auto_style(source: &[u8]) -> Vec<u8> {
    let cfb = CompoundFile::open(source).unwrap();
    let mut word = cfb.stream("WordDocument").unwrap();
    let mut table = cfb.stream("0Table").unwrap();
    let old = u32::from_le_bytes(word[0xa2..0xa6].try_into().unwrap()) as usize;
    let header = table[old..old + 20].to_vec();
    let mut normal = vec![0; 14];
    normal[2..4].copy_from_slice(&0xfff1u16.to_le_bytes());
    normal[4..6].copy_from_slice(&2u16.to_le_bytes());
    // UpxPapx istd=0, stored margins 5/7pt, both automatic flags enabled.
    let papx = [
        0, 0, 0x13, 0xa4, 100, 0, 0x14, 0xa4, 140, 0, 0x5b, 0x24, 1, 0x5c, 0x24, 1,
    ];
    normal.extend((papx.len() as u16).to_le_bytes());
    normal.extend(papx);
    normal.extend([4, 0, 0x43, 0x4a, 22, 0]); // paragraph base/mark 11pt
    let offset = table.len();
    table.extend(header);
    table.extend((normal.len() as u16).to_le_bytes());
    table.extend(normal);
    for _ in 1..15 {
        table.extend([0, 0]);
    }
    word[0xa2..0xa6].copy_from_slice(&(offset as u32).to_le_bytes());
    word[0xa6..0xaa].copy_from_slice(&((table.len() - offset) as u32).to_le_bytes());
    build_scoped_cfb(&[("WordDocument", word), ("0Table", table)])
}

fn document(bytes: &[u8]) -> Document {
    crate::doc::direct_model(&CompoundFile::open(bytes).unwrap(), 1024 * 1024)
        .unwrap()
        .document
}

#[test]
fn automatic_spacing_preserves_inherited_true_and_independent_direct_false() {
    let parts = ["inherit\r", "before\r", "after\r", "both\r"];
    let mut cp = 0;
    let runs: Vec<_> = parts
        .iter()
        .enumerate()
        .map(|(i, text)| {
            let start = cp;
            cp += text.len();
            let mut sprms = match i {
                0 => vec![],
                1 => vec![0x5b, 0x24, 0],
                2 => vec![0x5c, 0x24, 0],
                _ => vec![0x5b, 0x24, 0, 0x5c, 0x24, 0],
            };
            if !sprms.is_empty() && sprms.len() % 2 == 0 {
                sprms.extend([0x05, 0x24, 0]);
            }
            (start, cp, sprms)
        })
        .collect();
    let bytes = with_papx(&with_auto_style(&base(&parts.concat())), &runs);
    let doc = document(&bytes);
    let values: Vec<_> = doc
        .body
        .iter()
        .map(|element| {
            let BodyElement::Paragraph(p) = element else {
                panic!("paragraph")
            };
            assert_eq!(
                (p.space_before, p.space_after, p.default_font_size),
                (5.0, 7.0, Some(11.0))
            );
            (p.before_autospacing, p.after_autospacing)
        })
        .collect();
    assert_eq!(
        values,
        [
            (Some(true), Some(true)),
            (Some(false), Some(true)),
            (Some(true), Some(false)),
            (Some(false), Some(false))
        ]
    );
}

#[test]
fn automatic_spacing_piece_false_clears_style_and_direct_true() {
    let bytes = with_papx(
        &with_auto_style(&base("piece\r")),
        &[(0, 6, vec![0x5b, 0x24, 1])],
    );
    let cfb = CompoundFile::open(&bytes).unwrap();
    let mut word = cfb.stream("WordDocument").unwrap();
    let mut table = cfb.stream("0Table").unwrap();
    let offset = u32::from_le_bytes(word[0x1a2..0x1a6].try_into().unwrap()) as usize;
    let clx = table[offset..offset + 21].to_vec();
    let replacement = table.len();
    table.extend([1, 6, 0, 0x5b, 0x24, 0, 0x5c, 0x24, 0]);
    table.extend(clx);
    table[replacement + 28..replacement + 30].copy_from_slice(&1u16.to_le_bytes());
    word[0x1a2..0x1a6].copy_from_slice(&(replacement as u32).to_le_bytes());
    word[0x1a6..0x1aa].copy_from_slice(&30u32.to_le_bytes());
    let doc = document(&build_scoped_cfb(&[
        ("WordDocument", word),
        ("0Table", table),
    ]));
    let BodyElement::Paragraph(p) = &doc.body[0] else {
        panic!("paragraph")
    };
    assert_eq!(
        (p.before_autospacing, p.after_autospacing),
        (Some(false), Some(false))
    );
    assert_eq!((p.space_before, p.space_after), (5.0, 7.0));
}

#[test]
fn automatic_spacing_native_body_header_footer_and_note_keep_their_own_paragraphs() {
    let mut dop = default_note_dop();
    dop.resize(544, 0);
    dop[512..516].copy_from_slice(&4u32.to_le_bytes());
    let notes = NotesFixture {
        notes: &["\u{2}Note\r"],
        references: &[(1, true)],
        separators: ["\u{3}\r\r", "\u{4}\r\r", "", "", "", ""],
        dop: Some(dop),
    };
    let bytes = with_auto_style(&passive_special_source(&source_with_stories(
        "B\u{2}\r",
        &[(3, 2, 12240, 15840, 1, 720)],
        None,
        None,
        None,
        Some(&[None, Some("Header\r"), None, Some("Footer\r"), None, None]),
        Some(&notes),
    )));
    let doc = document(&bytes);
    assert_eq!(
        doc.settings
            .as_ref()
            .unwrap()
            .do_not_use_html_paragraph_auto_spacing,
        Some(true)
    );
    let paragraphs = [
        &doc.body[0],
        &doc.headers.default.as_ref().unwrap().body[0],
        &doc.footers.default.as_ref().unwrap().body[0],
        &doc.footnotes[0].content[0],
    ];
    let text: Vec<_> = paragraphs
        .iter()
        .map(|element| {
            let BodyElement::Paragraph(p) = element else {
                panic!("paragraph")
            };
            assert_eq!(
                (p.before_autospacing, p.after_autospacing),
                (Some(true), Some(true))
            );
            assert_eq!(
                (p.space_before, p.space_after, p.default_font_size),
                (5.0, 7.0, Some(11.0))
            );
            p.runs
                .iter()
                .filter_map(|run| match run {
                    docx_model::DocRun::Text(t) if t.note_ref.is_none() => Some(t.text.as_str()),
                    _ => None,
                })
                .collect::<String>()
        })
        .collect();
    assert_eq!(text, ["B", "Header", "Footer", "Note"]);
}

#[test]
fn automatic_spacing_dop_setting_reaches_canonical_document_without_neighbor_bit_aliasing() {
    for (size, flags, expected) in [
        (84, 0, None),
        (500, 0, None),
        (544, 0, Some(false)),
        (544, 4, Some(true)),
        (544, 8, Some(false)),
        (694, 12, Some(true)),
    ] {
        let bytes = base("settings\r");
        let cfb = CompoundFile::open(&bytes).unwrap();
        let mut word = cfb.stream("WordDocument").unwrap();
        let mut table = cfb.stream("0Table").unwrap();
        let offset = table.len();
        let mut dop = vec![0; size];
        dop[10..12].copy_from_slice(&720u16.to_le_bytes());
        if size >= 544 {
            dop[512..516].copy_from_slice(&(flags as u32).to_le_bytes());
        }
        table.extend(dop);
        word[0x192..0x196].copy_from_slice(&(offset as u32).to_le_bytes());
        word[0x196..0x19a].copy_from_slice(&(size as u32).to_le_bytes());
        let doc = document(&build_scoped_cfb(&[
            ("WordDocument", word.clone()),
            ("0Table", table.clone()),
        ]));
        let settings = doc.settings.unwrap();
        assert_eq!(
            settings.do_not_use_html_paragraph_auto_spacing, expected,
            "DOP {size}, flags {flags}"
        );
        assert_eq!(settings.adjust_line_height_in_table, Some(flags & 8 == 0));
        // A declared complete DOP must not read bit2 from a truncated payload.
        table.pop();
        let truncated = build_scoped_cfb(&[("WordDocument", word), ("0Table", table)]);
        assert!(
            crate::doc::direct_model(&CompoundFile::open(&truncated).unwrap(), 1024 * 1024)
                .unwrap_err()
                .contains("truncated Word document properties")
        );
    }
    assert!(document(&base("absent\r")).settings.is_none());
}
