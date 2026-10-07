//! Direct projection of footnotes and endnotes onto the DOCX model's notes.
//!
//! MS-DOC 2.3.2/2.3.5 (note documents), 2.8.16/2.8.17 (PlcffndRef/Txt),
//! 2.8.19/2.8.20 (PlcfendRef/Txt), 2.3.3 (separator stories), 2.7.2/2.7.4
//! (DOP note properties) and 2.6.4 (section note properties); ECMA-376
//! Part 1 17.11.
//!
//! The DOCX parser represents a note as a `DocxNote` whose blocks are parsed
//! like any story, a body `footnoteReference` as a superscript `TextRun`
//! tagged `NoteRef { kind, id }`, and the in-note `footnoteRef` as the same run
//! with an empty id. The direct model produces exactly that shape.
//!
//! The shared renderer numbers notes in first-reference order with the
//! document-wide numbering format and start (ECMA-376 17.11.17/.18/.20), lays
//! footnotes out at the page bottom and endnotes at the end of the document.
//! Native reserved stories retain explicit rule kinds and effective PAPX/CHPX
//! for its metric consumer. Numeric short-rule span remains library policy,
//! without a claim of Word-exact width or placement.
//!
//! Accepted note properties require automatic marks, continuous numbering
//! with one format/start per note kind (Arabic, Roman or letters), bottom-of-
//! page footnotes, document-end endnotes and the separator classes gated below.
//! Other properties stay rejected; literal custom marks need a separate shared
//! contract for marks that do not consume a number.

use super::{fields::StoryFields, story, ModelBudget};
use crate::doc::{
    formatting, header_fields, headers, notes, numbering, pictures, sections, settings,
    tokenize_with_fields, unsupported, Fields, Paragraph, Token,
};
use docx_model::{
    DocxNote, NativeNoteSeparatorStoriesWire, NativeNoteSeparatorsWire, NoteLayoutSettingsWire,
};

mod separators;

/// Last effective nFib whose note properties live in the DOP (MS-DOC 2.7.2).
const DOP_NOTE_PROPERTIES_MAX_NFIB: u16 = 0x00d9;

/// Validate native note properties against the implemented shared contract and
/// retain numbering/separator facts. ECMA-376 17.11.17/.18/.20 governs automatic
/// numbering; rule geometry and continuation scope follow the library policies
/// documented beside their gates and layout consumer.
// The parameters are independent note, header, section and acquisition facts
// plus the two mutable ownership budgets (formatting and model), passed
// individually like `story::project`'s projections.
#[allow(clippy::too_many_arguments)]
pub(super) fn validate(
    stories: &[Option<notes::Notes<'_>>],
    references: &notes::References,
    effective_nfib: u16,
    dop: Option<&settings::Properties>,
    sections: &[sections::Section],
    headers: Option<&headers::Headers<'_>>,
    formatting: &mut formatting::Formatting<'_>,
    budget: &mut ModelBudget,
) -> Result<Option<NoteLayoutSettingsWire>, String> {
    let present = |kind| {
        stories
            .iter()
            .flatten()
            .any(|notes| notes.kind == kind && !notes.entries.is_empty())
    };
    let (footnotes, endnotes) = (
        present(notes::Kind::Footnote),
        present(notes::Kind::Endnote),
    );
    if !footnotes && !endnotes {
        return Ok(None);
    }
    if references.iter().any(notes::Reference::custom) {
        return Err(unsupported(
            "Word custom note reference marks are not supported",
        ));
    }
    let dop = dop
        .map(|dop| dop.notes)
        .ok_or_else(|| unsupported("Word notes require document note properties"))?;
    // (format, start) for each kind: MSONFC and the first automatic number.
    let (footnote_numbering, endnote_numbering);
    if effective_nfib <= DOP_NOTE_PROPERTIES_MAX_NFIB {
        // MS-DOC 2.7.2/2.7.4: these documents keep note numbering and footnote
        // placement in the DOP. Section note SPRMs would be ambiguous.
        let formats = dop
            .formats
            .ok_or_else(|| unsupported("Word note number formats are missing"))?;
        if sections
            .iter()
            .any(|section| section.note_properties() != Default::default())
        {
            return Err(unsupported(
                "Word section note properties in a DOP-scoped document",
            ));
        }
        if footnotes && (dop.footnote_position != 1 || dop.footnote_restart != 0) {
            return Err(unsupported(footnote_message()));
        }
        if endnotes && dop.endnote_restart != 0 {
            return Err(unsupported(endnote_message()));
        }
        footnote_numbering = (formats.0, dop.footnote_start);
        endnote_numbering = (formats.1, dop.endnote_start);
    } else {
        // MS-DOC 2.6.4 defaults: fpcBottomPage, rncCont, no offset, Arabic
        // footnotes and lowercase-Roman endnotes. With continuous numbering
        // sprmSNFtn/sprmSNEdn add (value - 1) to every number of the section;
        // equal values in every section are a document-wide start value.
        let mut footnote_values = None;
        let mut endnote_values = None;
        for section in sections {
            let properties = section.note_properties();
            if footnotes
                && (!matches!(properties.footnote_position, None | Some(1))
                    || !matches!(properties.footnote_restart, None | Some(0)))
            {
                return Err(unsupported(footnote_message()));
            }
            if endnotes && !matches!(properties.endnote_restart, None | Some(0)) {
                return Err(unsupported(endnote_message()));
            }
            let footnote = (
                properties.footnote_format.unwrap_or(0),
                properties.footnote_offset.unwrap_or(1),
            );
            let endnote = (
                properties.endnote_format.unwrap_or(2),
                properties.endnote_offset.unwrap_or(1),
            );
            if *footnote_values.get_or_insert(footnote) != footnote && footnotes {
                return Err(unsupported(footnote_message()));
            }
            if *endnote_values.get_or_insert(endnote) != endnote && endnotes {
                return Err(unsupported(endnote_message()));
            }
        }
        footnote_numbering = footnote_values.unwrap_or((0, 1));
        endnote_numbering = endnote_values.unwrap_or((2, 1));
    }
    let mut settings = NoteLayoutSettingsWire {
        footnote_number_format: footnotes
            .then(|| number_format(footnote_numbering.0, footnote_message()))
            .transpose()?,
        footnote_number_start: footnotes
            .then(|| number_start(footnote_numbering.1, footnote_message()))
            .transpose()?,
        endnote_number_format: endnotes
            .then(|| number_format(endnote_numbering.0, endnote_message()))
            .transpose()?,
        endnote_number_start: endnotes
            .then(|| number_start(endnote_numbering.1, endnote_message()))
            .transpose()?,
        ..NoteLayoutSettingsWire::default()
    };
    // DopBase.epc is not scoped by nFib. Only end-of-document placement (3)
    // matches the shared layout; sprmSFEndnote is then irrelevant.
    if endnotes && dop.endnote_position != 3 {
        return Err(unsupported(
            "Word end-of-section endnote placement is not supported",
        ));
    }
    // MS-DOC 2.3.3/2.8.22 assigns story roles; 2.6.1 sprmCFSpec assigns
    // U+0003/U+0004 short/full semantics independently of those roles. Retain
    // actual PAPX/CHPX and native provenance for the shared metric consumer.
    // Footnote short/short is supported within its closed metric-only contract.
    // Numeric short-rule span remains the existing library layout policy;
    // admission does not claim Office-exact rule placement or width.
    let headers = headers
        .ok_or_else(|| unsupported("Word notes without separator stories are not supported"))?;
    let mut native = NativeNoteSeparatorsWire::default();
    for (present, base) in [(footnotes, 0), (endnotes, 3)] {
        if !present {
            continue;
        }
        let ordinary = separators::acquire(headers, base, formatting, budget)?;
        let continuation = separators::acquire(headers, base + 1, formatting, budget)?;
        // A paragraph-only notice owns real PAPX/CHPX too. Acquiring it does
        // not equate its formatting with DOCX absence or a default template.
        let notice = separators::acquire(headers, base + 2, formatting, budget)?;
        // Only the footnote paginator consumes continuation roles. Preserve
        // the existing endnote envelope until endnote continuation is wired.
        // Apply metric eligibility only to this newly admitted class: existing
        // short/full admission keeps its previous producer behavior.
        let short_continuation = base == 0
            && continuation.standard_rule(separators::RuleKind::Short)
            && [&ordinary, &continuation, &notice]
                .iter()
                .all(|story| story.metric_only());
        if !ordinary.standard_rule(separators::RuleKind::Short)
            || !(continuation.standard_rule(separators::RuleKind::Full) || short_continuation)
            || !matches!(headers.separator_text(base + 2), "" | "\r\r")
        {
            return Err(unsupported("custom Word note separators are not supported"));
        }
        // A rule retains its explicit kind and effective character metrics.
        // Continuation remains caller opt-in; no document/default option changes.
        let slot = if base == 0 {
            settings.footnote_separator = ordinary.mark();
            settings.footnote_continuation_separator = continuation.mark();
            &mut native.footnote
        } else {
            settings.endnote_separator = ordinary.mark();
            &mut native.endnote
        };
        *slot = Some(NativeNoteSeparatorStoriesWire {
            separator: ordinary.into_wire()?,
            continuation_separator: continuation.into_wire()?,
            continuation_notice: notice.into_wire()?,
        });
    }
    // One new retained container per document. Moved payloads keep their
    // cumulative acquisition charges; nothing in it is a second copy.
    budget.charge(std::mem::size_of::<NativeNoteSeparatorsWire>())?;
    settings.native_separators = Some(Box::new(native));
    Ok(Some(settings))
}

/// MS-OSHARED 2.2.1.3 MSONFC values whose ECMA-376 17.18.59 format the shared
/// note renderer formats natively: Arabic, upper/lower Roman and letters.
fn number_format(value: u16, message: &'static str) -> Result<String, String> {
    match value {
        0 => Ok("decimal"),
        1 => Ok("upperRoman"),
        2 => Ok("lowerRoman"),
        3 => Ok("upperLetter"),
        4 => Ok("lowerLetter"),
        _ => Err(unsupported(message)),
    }
    .map(str::to_string)
}

/// A first automatic number of at least 1 (the 14-bit DOP value or the
/// section SPRM value, at most 16383).
fn number_start(value: u16, message: &'static str) -> Result<i64, String> {
    if value == 0 {
        return Err(unsupported(message));
    }
    Ok(i64::from(value))
}

fn footnote_message() -> &'static str {
    "Word footnote numbering or placement other than continuous Roman, letter or Arabic numbering at the page bottom is not supported"
}

fn endnote_message() -> &'static str {
    "Word endnote numbering other than continuous Roman, letter or Arabic numbering is not supported"
}

/// Turn main-story automatic note characters into references. Every
/// character must name a reference and every reference must be displayed.
pub(super) fn restore_references(
    references: &notes::References,
    paragraphs: &mut [Paragraph],
    placed: &mut usize,
) -> Result<(), String> {
    for paragraph in paragraphs {
        for (token, cp) in &mut paragraph.tokens {
            let token = match token {
                Token::Linked(linked) => &mut linked.token,
                token => token,
            };
            if matches!(token, Token::NoteMarker) {
                let reference = references.get(*cp).ok_or_else(|| {
                    unsupported("Word automatic note character without a note reference")
                })?;
                *token = Token::NoteReference(reference.clone());
                *placed += 1;
            }
        }
    }
    Ok(())
}

pub(super) fn check_all_references_placed(
    references: &notes::References,
    placed: usize,
) -> Result<(), String> {
    if placed != references.iter().count() {
        return Err(unsupported("Word note reference is not displayed"));
    }
    Ok(())
}

/// Project every note of one note document, in PLC order, as DOCX notes whose
/// ids match the main-story references (`Reference::id`).
pub(super) fn project(
    notes: &notes::Notes<'_>,
    table: &header_fields::Table,
    formatting: &mut formatting::Formatting<'_>,
    pictures: &mut pictures::Store<'_>,
    budget: &mut ModelBudget,
    table_sequence: &mut usize,
    output: &mut Vec<DocxNote>,
) -> Result<(), String> {
    let mut partitions = Vec::with_capacity(notes.entries.len() + 1);
    let mut end = 0;
    for entry in &notes.entries {
        partitions.push(entry.cp);
        end = entry.cp + notes.story.text[entry.text.clone()].encode_utf16().count();
    }
    partitions.push(end);
    let fields = StoryFields::analyze(&notes.story.text, table, &partitions)?;
    for (index, entry) in notes.entries.iter().enumerate() {
        let text = &notes.story.text[entry.text.clone()];
        budget.charge(text.len())?;
        let mut paragraphs = tokenize_with_fields(text, &mut Fields::default(), entry.cp, true);
        fields.apply(entry.cp, &mut paragraphs)?;
        for paragraph in &mut paragraphs {
            for (token, _) in &mut paragraph.tokens {
                let token = match token {
                    Token::Linked(linked) => &mut linked.token,
                    token => token,
                };
                if matches!(token, Token::NoteMarker) {
                    // `validate` rejected custom marks, so every note is
                    // automatic; each U+0002 shows the enclosing note number.
                    *token = Token::NoteNumber(notes.kind);
                }
            }
        }
        // Lists restart in every note.
        let mut numbering = numbering::direct::Store::default();
        numbering.begin_story()?;
        let mut content = Vec::new();
        story::project(
            &notes.story,
            paragraphs,
            formatting,
            &mut numbering,
            pictures,
            None,
            budget,
            &mut content,
            None,
            table_sequence,
        )?;
        let id = (index + 1).to_string();
        budget.charge(id.capacity())?;
        budget.push(output, DocxNote { id, content })?;
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::super::tests::{
        default_note_dop, passive_special_source, source_with_stories, NotesFixture,
    };
    use crate::cfb::CompoundFile;
    use docx_model::{BodyElement, DocRun, Document};

    const SEPARATORS: [&str; 6] = ["\u{3}\r\r", "\u{4}\r\r", "", "", "", ""];

    fn document(text: &str, fixture: &NotesFixture<'_>) -> Result<Document, String> {
        let sections = [(text.encode_utf16().count(), 2, 12_240, 15_840, 1, 720)];
        let slots = [None; 6];
        let bytes = source_with_stories(
            text,
            &sections,
            None,
            None,
            None,
            Some(&slots),
            Some(fixture),
        );
        // MS-DOC 2.3.2: note references and numbers are special characters.
        let bytes = passive_special_source(&bytes);
        super::super::super::direct_model(&CompoundFile::open(&bytes).unwrap(), 1024 * 1024)
            .map(|result| result.document)
    }

    fn runs(element: &BodyElement) -> Vec<String> {
        let BodyElement::Paragraph(paragraph) = element else {
            panic!("paragraph expected");
        };
        paragraph
            .runs
            .iter()
            .map(|run| match run {
                DocRun::Text(text) => match &text.note_ref {
                    Some(note) => format!(
                        "{}:{}:{}:{}",
                        note.kind,
                        note.id,
                        text.text,
                        text.vert_align.as_deref().unwrap_or_default()
                    ),
                    None => text.text.clone(),
                },
                DocRun::Field(field) => format!("field:{}", field.field_type),
                _ => "other".into(),
            })
            .collect()
    }

    fn fixture<'a>(notes: &'a [&'a str], references: &'a [(u32, bool)]) -> NotesFixture<'a> {
        NotesFixture {
            notes,
            references,
            separators: SEPARATORS,
            dop: Some(default_note_dop()),
        }
    }

    #[test]
    fn separator_acquisition_resolves_special_and_formatting_before_role_policy() {
        let mut fixture = fixture(&["\u{2} note\r"], &[(1, true)]);
        fixture.separators[1] = "\u{3}\r\r";
        fixture.separators[2] = "\r\r";
        let bytes = source_with_stories(
            "A\u{2}\r",
            &[(3, 2, 12_240, 15_840, 1, 720)],
            None,
            Some(24),
            None,
            Some(&[None; 6]),
            Some(&fixture),
        );
        let bytes = passive_special_source(&bytes);
        crate::doc::with_acquired_doc(&CompoundFile::open(&bytes).unwrap(), |mut facts| {
            let headers = facts.headers.as_ref().unwrap();
            let mut budget = super::super::ModelBudget::new(1024 * 1024);
            for slot in [0, 1] {
                let acquired =
                    super::separators::acquire(headers, slot, &mut facts.formatting, &mut budget)?;
                let rule = acquired.rule.as_ref().unwrap();
                assert_eq!(rule.kind, Some(super::separators::RuleKind::Short));
                assert!(rule.character.picture.passive_special());
                assert_eq!(rule.character.run.as_ref().unwrap().font_size, 12.0);
                let paragraph = acquired.paragraph.as_ref().unwrap();
                assert_eq!(paragraph.paragraph.default_font_size, Some(12.0));
                assert_eq!(
                    paragraph
                        .mark
                        .as_ref()
                        .unwrap()
                        .run
                        .as_ref()
                        .unwrap()
                        .font_size,
                    12.0
                );
                assert_eq!(acquired.source.content, "\u{3}\r");
                assert_eq!(acquired.source.cp, (slot * 3)..(slot * 3 + 3));
            }
            let notice =
                super::separators::acquire(headers, 2, &mut facts.formatting, &mut budget)?;
            assert!(notice.rule.is_none());
            assert_eq!(notice.source.content_cp, 6..7);
            assert_eq!(
                notice
                    .paragraph
                    .as_ref()
                    .unwrap()
                    .mark
                    .as_ref()
                    .unwrap()
                    .run
                    .as_ref()
                    .unwrap()
                    .font_size,
                12.0
            );
            // Admission preserves the acquired short continuation and its
            // formatting rather than substituting the continuation-role default.
            let settings = super::validate(
                &facts.note_stories,
                &facts.note_references,
                facts.effective_nfib,
                facts.document_settings.as_ref(),
                &facts.sections,
                Some(headers),
                &mut facts.formatting,
                &mut budget,
            )
            .unwrap()
            .unwrap();
            assert_eq!(
                settings.footnote_continuation_separator,
                Some(docx_model::NoteSeparatorMark::Short)
            );
            let stories = settings.native_separators.unwrap().footnote.unwrap();
            let continuation = stories.continuation_separator.rule.unwrap();
            assert_eq!(
                continuation.mark,
                Some(docx_model::NoteSeparatorMark::Short)
            );
            assert_eq!(continuation.control.run.unwrap().font_size, 12.0);
            let notice = stories.continuation_notice.paragraph.unwrap();
            assert_eq!(notice.content_mark.unwrap().run.unwrap().font_size, 12.0);
            Ok(())
        })
        .unwrap();
    }

    /// Production native bytes whose Normal style is 14pt, so every retained
    /// CHPX/PAPX font size comes from the real style cascade, not a default.
    fn styled_source(notice: &'static str) -> Vec<u8> {
        let mut fixture = fixture(&["\u{2} note\r"], &[(1, true)]);
        fixture.separators[2] = notice;
        let bytes = source_with_stories(
            "A\u{2}\r",
            &[(3, 2, 12_240, 15_840, 1, 720)],
            None,
            Some(28),
            None,
            Some(&[None; 6]),
            Some(&fixture),
        );
        passive_special_source(&bytes)
    }

    #[test]
    fn admitted_separator_formatting_moves_into_the_private_wire() {
        use docx_model::{NativeNoteSeparatorStoryClass as Class, NoteSeparatorMark as Mark};
        for notice in ["", "\r\r"] {
            let bytes = styled_source(notice);
            let cfb = CompoundFile::open(&bytes).unwrap();
            let result = super::super::super::direct_model(&cfb, 1024 * 1024).unwrap();
            let settings = result.document.note_layout_settings.unwrap();
            // Retention leaves the scalar marks that select the drawn rule.
            assert_eq!(settings.footnote_separator, Some(Mark::Short));
            assert_eq!(settings.footnote_continuation_separator, Some(Mark::Full));
            let native = settings.native_separators.unwrap();
            assert!(native.endnote.is_none());
            let stories = native.footnote.unwrap();
            for (story, mark, start) in [
                (&stories.separator, Mark::Short, 0),
                (&stories.continuation_separator, Mark::Full, 3),
            ] {
                assert_eq!(story.class, Class::Rule);
                assert_eq!(
                    (story.content_start_cp, story.content_end_cp, story.guard_cp),
                    (start, start + 2, Some(start + 2))
                );
                let rule = story.rule.as_ref().unwrap();
                assert_eq!(rule.mark, Some(mark));
                assert_eq!(rule.source.header_cp, start);
                assert_eq!(rule.control.run.as_ref().unwrap().font_size, 14.0);
                // The content paragraph mark is its own character, not the
                // control and not the guard.
                let paragraph = story.paragraph.as_ref().unwrap();
                assert_eq!(paragraph.source.header_cp, start + 1);
                assert_ne!(paragraph.source.fc, rule.source.fc);
                assert_eq!(paragraph.paragraph.default_font_size, Some(14.0));
                let content_mark = paragraph.content_mark.as_ref().unwrap();
                assert_eq!(content_mark.run.as_ref().unwrap().font_size, 14.0);
            }
            let notice_story = &stories.continuation_notice;
            assert!(notice_story.rule.is_none());
            if notice.is_empty() {
                assert_eq!(notice_story.class, Class::Empty);
                assert!(notice_story.guard_cp.is_none() && notice_story.paragraph.is_none());
            } else {
                assert_eq!(notice_story.class, Class::ParagraphOnly);
                assert_eq!(notice_story.guard_cp, Some(7));
                let paragraph = notice_story.paragraph.as_ref().unwrap();
                assert_eq!(paragraph.source.header_cp, 6);
                let content_mark = paragraph.content_mark.as_ref().unwrap();
                assert_eq!(content_mark.run.as_ref().unwrap().font_size, 14.0);
            }
        }
    }

    #[test]
    fn retained_separator_container_is_charged_once_without_payload_copies() {
        let bytes = styled_source("\r\r");
        let cfb = CompoundFile::open(&bytes).unwrap();
        let limit = 1024 * 1024;
        let acquired = crate::doc::with_acquired_doc(&cfb, |mut facts| {
            let headers = facts.headers.as_ref().unwrap();
            let mut budget = super::super::ModelBudget::new(limit);
            for slot in 0..3 {
                super::separators::acquire(headers, slot, &mut facts.formatting, &mut budget)?;
            }
            Ok(limit - budget.remaining_bytes)
        })
        .unwrap();
        let validate = |limit| {
            crate::doc::with_acquired_doc(&cfb, |mut facts| {
                let mut budget = super::super::ModelBudget::new(limit);
                let result = super::validate(
                    &facts.note_stories,
                    &facts.note_references,
                    facts.effective_nfib,
                    facts.document_settings.as_ref(),
                    &facts.sections,
                    facts.headers.as_ref(),
                    &mut facts.formatting,
                    &mut budget,
                );
                Ok((result, limit - budget.remaining_bytes))
            })
            .unwrap()
        };
        let (result, used) = validate(limit);
        assert!(result.unwrap().unwrap().native_separators.is_some());
        // Moved paragraph/run payloads were charged by acquisition; only the
        // one boxed container is new.
        let container = std::mem::size_of::<docx_model::NativeNoteSeparatorsWire>();
        assert_eq!(used, acquired + container);
        assert_eq!(validate(used - 1).0.unwrap_err(), "OUTPUT_TOO_LARGE");
    }

    #[test]
    fn raw_separator_controls_without_special_meaning_stay_rejected() {
        let fixture = fixture(&["\u{2} note\r"], &[(1, true)]);
        let bytes = source_with_stories(
            "A\u{2}\r",
            &[(3, 2, 12_240, 15_840, 1, 720)],
            None,
            None,
            None,
            Some(&[None; 6]),
            Some(&fixture),
        );
        let bytes = passive_special_source(&bytes);
        let cfb = CompoundFile::open(&bytes).unwrap();
        let mut word = cfb.stream("WordDocument").unwrap();
        let table = cfb.stream("0Table").unwrap();
        let bte = u32::from_le_bytes(word[0xfa..0xfe].try_into().unwrap()) as usize;
        let page_number = u32::from_le_bytes(table[bte + 8..bte + 12].try_into().unwrap()) as usize;
        let header_cp = u32::from_le_bytes(word[0x4c..0x50].try_into().unwrap())
            + u32::from_le_bytes(word[0x50..0x54].try_into().unwrap());
        let page = &mut word[page_number * 512..(page_number + 1) * 512];
        let start_fc = u32::from_le_bytes(page[..4].try_into().unwrap());
        let end_fc = u32::from_le_bytes(page[4..8].try_into().unwrap());
        page[4..8].copy_from_slice(&(start_fc + header_cp * 2).to_le_bytes());
        page[8..12].copy_from_slice(&end_fc.to_le_bytes());
        page[12..14].copy_from_slice(&[32, 34]);
        page[64..68].copy_from_slice(&[3, 0x55, 0x08, 1]);
        page[68..72].copy_from_slice(&[3, 0x55, 0x08, 0]);
        page[511] = 2;
        let bytes = crate::cfb::test_support::build_scoped_cfb(&[
            ("WordDocument", word),
            ("0Table", table),
        ]);
        assert!(
            crate::doc::direct_model(&CompoundFile::open(&bytes).unwrap(), 1024 * 1024)
                .unwrap_err()
                .contains("separators")
        );
    }

    #[test]
    fn footnotes_project_onto_the_docx_note_model() {
        let notes = ["\u{2} one\r", "\u{2} two\rsecond\r"];
        let references = [(1, true), (3, true)];
        let document = document("A\u{2}B\u{2}\r", &fixture(&notes, &references)).unwrap();
        let settings = document.note_layout_settings.as_ref().unwrap();
        assert_eq!(
            settings.footnote_separator,
            Some(docx_model::NoteSeparatorMark::Short)
        );
        assert_eq!(
            settings.footnote_continuation_separator,
            Some(docx_model::NoteSeparatorMark::Full)
        );
        assert_eq!(settings.endnote_separator, None);
        assert_eq!(
            runs(&document.body[0]),
            ["A", "footnote:1:1:super", "B", "footnote:2:2:super"]
        );
        assert!(document.endnotes.is_empty());
        let ids: Vec<_> = document
            .footnotes
            .iter()
            .map(|note| note.id.as_str())
            .collect();
        assert_eq!(ids, ["1", "2"]);
        assert_eq!(
            runs(&document.footnotes[0].content[0]),
            ["footnote:::super", " one"]
        );
        assert_eq!(document.footnotes[1].content.len(), 2);
        assert_eq!(runs(&document.footnotes[1].content[1]), ["second"]);
    }

    #[test]
    fn note_fields_use_the_footnote_field_table() {
        let notes = ["\u{2} p\u{13}PAGE\u{14}9\u{15}\r"];
        let references = [(1, true)];
        let document = document("A\u{2}\r", &fixture(&notes, &references)).unwrap();
        assert_eq!(
            runs(&document.footnotes[0].content[0]),
            ["footnote:::super", " p", "field:page"]
        );
    }

    #[test]
    fn unrepresentable_note_properties_are_rejected() {
        let notes = ["\u{2} one\r"];
        let references = [(1, true)];
        let reject = |fixture: NotesFixture<'_>, text: &str| document(text, &fixture).unwrap_err();
        // A literal custom mark would still be numbered by the renderer.
        let custom = [(1, false)];
        assert!(reject(fixture(&notes, &custom), "A*\r").contains("custom note reference"));
        for (offset, value, message) in [
            // Chicago numbering, per-section and per-page restarts, a zero
            // start and beneath-text placement have no shared rendering.
            (492usize, 9u8, "footnote numbering"),
            (2, 1 | (1 << 2), "footnote numbering"),
            (2, 2 | (1 << 2), "footnote numbering"),
            (2, 0, "footnote numbering"),
            (0, 2 << 5, "footnote numbering"),
        ] {
            let mut properties = fixture(&notes, &references);
            let dop = properties.dop.as_mut().unwrap();
            if offset == 2 {
                dop[2..4].copy_from_slice(&u16::from(value).to_le_bytes());
            } else {
                dop[offset] = value;
            }
            assert!(reject(properties, "A\u{2}\r").contains(message), "{offset}");
        }
        let mut missing = fixture(&notes, &references);
        missing.dop = None;
        assert!(reject(missing, "A\u{2}\r").contains("note properties"));
        for separators in [
            ["", "\u{4}\r\r", "", "", "", ""],
            ["\u{3}\r\r\r", "\u{4}\r\r", "", "", "", ""],
            ["\u{4}\r\r", "\u{3}\r\r", "", "", "", ""],
            ["\u{3}\r\r", "\u{4}\r\r", "x\r\r", "", "", ""],
        ] {
            let mut properties = fixture(&notes, &references);
            properties.separators = separators;
            assert!(reject(properties, "A\u{2}\r").contains("separators"));
        }
    }

    #[test]
    fn dop_note_format_and_start_become_document_note_numbering() {
        let notes = ["\u{2} one\r"];
        let references = [(1, true)];
        let mut properties = fixture(&notes, &references);
        let dop = properties.dop.as_mut().unwrap();
        dop[492] = 2; // nfcFtnRef msonfcLCRoman
        dop[2..4].copy_from_slice(&(3u16 << 2).to_le_bytes()); // nFtn 3
        let settings = document("A\u{2}\r", &properties)
            .unwrap()
            .note_layout_settings
            .unwrap();
        assert_eq!(
            settings.footnote_number_format.as_deref(),
            Some("lowerRoman")
        );
        assert_eq!(settings.footnote_number_start, Some(3));
        assert_eq!(settings.endnote_number_format, None);
    }

    #[test]
    fn every_reference_must_be_displayed_exactly_once() {
        let notes = ["\u{2} one\r"];
        let references = [(5, true)];
        // The only U+0002 lies inside a hidden field instruction.
        let text = "A\u{13}IF \u{2}\u{15}\r";
        let error = document(text, &fixture(&notes, &references)).unwrap_err();
        assert!(error.contains("not displayed"), "{error}");
    }
}
