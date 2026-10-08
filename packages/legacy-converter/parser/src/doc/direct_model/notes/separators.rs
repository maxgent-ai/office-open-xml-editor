//! Native reserved-story acquisition, separate from shared note placement.
//! MS-DOC 2.3.3/2.8.22 define slots and guards; 2.6.1 sprmCFSpec gives U+0003
//! and U+0004 rule semantics only after the actual CHPX/style/PCD cascade.
//! Role does not select kind. Numeric rule geometry is a shared layout policy,
//! not acquired here. Supported CHPX/PAPX and native provenance move into the
//! private wire; the shared consumer uses visible rule/control mark metrics.
//! Acquisition alone does not admit custom content or establish Office-exact
//! width/placement fidelity.

use super::super::{payload, ModelBudget};
use crate::doc::{formatting, headers, unsupported};
use docx_model::{
    NativeNoteSeparatorCharacterWire, NativeNoteSeparatorParagraphWire,
    NativeNoteSeparatorRuleWire, NativeNoteSeparatorSourceWire, NativeNoteSeparatorStoryClass,
    NativeNoteSeparatorStoryWire, NoteSeparatorMark,
};

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum RuleKind {
    Short,
    Full,
}

/// CP is relative to the header document; FC and PRM retain its actual piece.
/// These native addresses must not be mistaken for shared model source paths.
pub(super) struct SourceAddress {
    pub(super) cp: usize,
    pub(super) fc: usize,
    pub(super) prm: u16,
    pub(super) paragraph_style: usize,
}

impl SourceAddress {
    fn wire(&self) -> NativeNoteSeparatorSourceWire {
        NativeNoteSeparatorSourceWire {
            header_cp: self.cp,
            fc: self.fc,
            prm: self.prm,
            paragraph_style: self.paragraph_style,
        }
    }
}

pub(super) struct RuleFacts {
    pub(super) kind: Option<RuleKind>,
    pub(super) character: formatting::DirectRuleControlFacts,
    pub(super) source: SourceAddress,
}

pub(super) struct ParagraphFacts {
    pub(super) paragraph: docx_model::DocParagraph,
    pub(super) mark: Option<formatting::DirectRuleControlFacts>,
    pub(super) source: SourceAddress,
    numbered: bool,
    framed: bool,
}

pub(super) struct AcquiredSeparator<'a> {
    pub(super) source: headers::SeparatorStory<'a>,
    pub(super) rule: Option<RuleFacts>,
    pub(super) paragraph: Option<ParagraphFacts>,
}

impl RuleKind {
    fn mark(self) -> NoteSeparatorMark {
        match self {
            Self::Short => NoteSeparatorMark::Short,
            Self::Full => NoteSeparatorMark::Full,
        }
    }
}

impl AcquiredSeparator<'_> {
    /// Preserve explicit native rule kind independently of its destination role.
    pub(super) fn mark(&self) -> Option<NoteSeparatorMark> {
        self.rule.as_ref()?.kind.map(RuleKind::mark)
    }

    /// Move the acquired facts of an admitted story into the private wire.
    /// Owned paragraph/run payloads move without a clone, so their acquisition
    /// charges stand; the caller charges the new retained container. Borrowed
    /// header text is never copied: CPs locate it.
    pub(super) fn into_wire(self) -> Result<NativeNoteSeparatorStoryWire, String> {
        let class = match (self.source.guard_cp, self.source.content) {
            (None, "") => NativeNoteSeparatorStoryClass::Empty,
            (Some(_), "") => NativeNoteSeparatorStoryClass::GuardOnly,
            (Some(_), "\r") => NativeNoteSeparatorStoryClass::ParagraphOnly,
            _ if self.rule.is_some() => NativeNoteSeparatorStoryClass::Rule,
            _ => return Err(unsupported("custom Word note separators are not supported")),
        };
        Ok(NativeNoteSeparatorStoryWire {
            class,
            content_start_cp: self.source.content_cp.start,
            content_end_cp: self.source.content_cp.end,
            guard_cp: self.source.guard_cp,
            rule: self.rule.map(|rule| NativeNoteSeparatorRuleWire {
                mark: rule.kind.map(RuleKind::mark),
                source: rule.source.wire(),
                control: NativeNoteSeparatorCharacterWire {
                    run: rule.character.run,
                    insertion: rule.character.insertion,
                },
            }),
            paragraph: self
                .paragraph
                .map(|paragraph| NativeNoteSeparatorParagraphWire {
                    source: paragraph.source.wire(),
                    paragraph: paragraph.paragraph,
                    content_mark: paragraph.mark.map(|mark| NativeNoteSeparatorCharacterWire {
                        run: mark.run,
                        insertion: mark.insertion,
                    }),
                    numbered: paragraph.numbered,
                    framed: paragraph.framed,
                }),
        })
    }

    /// The shared native separator consumer has metric participants, not
    /// authored text, numbering/frame cascades, or tracked-insertion payloads.
    /// sprmCSymbol (MS-DOC 2.6.1/2.9.47) can populate an otherwise empty run;
    /// check the acquired payload before admitting the new short/short class.
    /// Structure and passive-special visibility are checked by standard_rule.
    pub(super) fn metric_only(&self) -> bool {
        let character = |facts: &formatting::DirectRuleControlFacts| {
            !facts.insertion && facts.run.as_ref().is_none_or(|run| run.text.is_empty())
        };
        self.rule
            .as_ref()
            .is_none_or(|rule| character(&rule.character))
            && self.paragraph.as_ref().is_none_or(|paragraph| {
                !paragraph.numbered
                    && !paragraph.framed
                    && paragraph.mark.as_ref().is_some_and(character)
            })
    }

    pub(super) fn standard_rule(&self, kind: RuleKind) -> bool {
        let expected = match kind {
            RuleKind::Short => "\u{3}\r",
            RuleKind::Full => "\u{4}\r",
        };
        self.source.cp.end - self.source.cp.start == 3
            && self.source.guard_cp.is_some()
            && self.source.content == expected
            && self.rule.as_ref().is_some_and(|rule| {
                rule.source.cp == self.source.content_cp.start
                    && rule.kind == Some(kind)
                    && rule.character.picture.passive_special()
                    && rule.character.run.is_some()
                    && !rule.character.insertion
            })
            && self.paragraph.as_ref().is_some_and(|paragraph| {
                paragraph.source.cp + 1 == self.source.content_cp.end
                    && !paragraph.numbered
                    && !paragraph.framed
                    && paragraph.mark.as_ref().is_some_and(|mark| !mark.insertion)
            })
    }
}

pub(super) fn acquire<'a>(
    headers: &'a headers::Headers<'_>,
    slot: usize,
    formatting: &mut formatting::Formatting<'_>,
    budget: &mut ModelBudget,
) -> Result<AcquiredSeparator<'a>, String> {
    budget.charge(std::mem::size_of::<AcquiredSeparator<'_>>())?;
    let source = headers.separator(slot);
    let marker = match source.content {
        "\u{3}" | "\u{3}\r" => Some(RuleKind::Short),
        "\u{4}" | "\u{4}\r" => Some(RuleKind::Full),
        _ => None,
    };
    let mut paragraph_style = None;
    let paragraph = if marker.is_some() || source.content == "\r" {
        // A content paragraph mark owns paragraph formatting. A separator need
        // not contain a whole paragraph; for that distinct, currently gated
        // class query its control's paragraph range, never the guard's style.
        let paragraph_cp = if source.content.ends_with('\r') {
            source.content_cp.end - 1
        } else {
            source.content_cp.start
        };
        let (_, mark_fc, mark_piece) = headers
            .story
            .position(paragraph_cp)
            .ok_or_else(|| unsupported("Word separator paragraph outside header story"))?;
        let style = formatting.paragraph_style(mark_fc)?;
        paragraph_style = Some(style);
        let paragraph = formatting.direct_paragraph(
            style,
            None,
            mark_fc,
            mark_piece.prm,
            &headers.story.prcs,
        )?;
        // Numbering marker properties are a deferred, owned cascade. A rule
        // does not activate numbering: retain its presence, not that uncharged
        // marker payload, alongside the charged visible paragraph metadata.
        let numbered = paragraph.numbering.is_some();
        let framed = paragraph.frame_gap.is_some() || paragraph.native_frame.is_some();
        let paragraph = paragraph.paragraph;
        budget.paragraph(&paragraph)?;
        // Retain the complete supported CHPX of the content mark separately
        // from paragraph font facts. The terminal guard is never its substitute.
        // A partial rule has no authored content mark and keeps that distinction.
        let mark = if source.content.ends_with('\r') {
            let mark = formatting.direct_rule_control(
                style,
                mark_fc,
                mark_piece.prm,
                &headers.story.prcs,
            )?;
            if let Some(run) = &mark.run {
                budget.charge(payload::text_run(run)?)?;
            }
            Some(mark)
        } else {
            None
        };
        Some(ParagraphFacts {
            paragraph,
            mark,
            numbered,
            framed,
            source: SourceAddress {
                cp: paragraph_cp,
                fc: mark_fc,
                prm: mark_piece.prm,
                paragraph_style: style,
            },
        })
    } else {
        None
    };
    let rule = if let Some(marker) = marker {
        let style = paragraph_style
            .ok_or_else(|| unsupported("Word separator paragraph style is missing"))?;
        let (_, fc, piece) = headers
            .story
            .position(source.content_cp.start)
            .ok_or_else(|| unsupported("Word separator control outside header story"))?;
        let character =
            formatting.direct_rule_control(style, fc, piece.prm, &headers.story.prcs)?;
        if let Some(run) = &character.run {
            budget.charge(payload::text_run(run)?)?;
        }
        Some(RuleFacts {
            kind: character.picture.special.then_some(marker),
            character,
            source: SourceAddress {
                cp: source.content_cp.start,
                fc,
                prm: piece.prm,
                paragraph_style: style,
            },
        })
    } else {
        None
    };
    Ok(AcquiredSeparator {
        source,
        rule,
        paragraph,
    })
}
