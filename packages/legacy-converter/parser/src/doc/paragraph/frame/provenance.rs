//! Effective native frame operands and their winning cascade layers. These are
//! producer facts, separate from frame equality and canonical shared paths.

use super::{Frame, TableParagraphFrame};

/// MS-DOC 2.6.2 property axes. Native values retain twips/decoded enum bits.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(in crate::doc) enum FrameProperty {
    X,
    Y,
    Width,
    Height,
    Anchors,
    Wrap,
    HorizontalSpace,
    VerticalSpace,
    Locked,
    NoOverlap,
    DropCap,
    TextFlow,
}

impl FrameProperty {
    pub(in crate::doc) fn from_code(code: u16) -> Option<Self> {
        Some(match code {
            0x8418 => Self::X,
            0x8419 => Self::Y,
            0x841a => Self::Width,
            0x442b => Self::Height,
            0x261b => Self::Anchors,
            0x2423 => Self::Wrap,
            0x842f => Self::HorizontalSpace,
            0x842e => Self::VerticalSpace,
            0x2430 => Self::Locked,
            0x2462 => Self::NoOverlap,
            0x442c => Self::DropCap,
            0x443a => Self::TextFlow,
            _ => return None,
        })
    }
}

/// Winning layer, including explicit writes of default values. FC/PRM name
/// the paragraph mark's piece; style/list IDs belong to this DOC. Redirected
/// PAPX stays owned by its PAPX layer. Provenance is never a mirror predicate.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(in crate::doc) enum FrameOrigin {
    Decoded,
    ParagraphStyle(usize),
    Papx {
        fc: usize,
    },
    Piece {
        fc: usize,
        prm: u16,
    },
    LinkedListStyle(usize),
    ListLevel {
        instance: usize,
        list: usize,
        level: u8,
    },
}

/// Fixed inline metadata: no per-property heap, full history or ancestor copy.
/// Style origins copy through the bounded document-owned template cache, then
/// direct/piece/list writes replace individual winners before classification.
pub(in crate::doc) type FrameOrigins = [Option<FrameOrigin>; 12];

#[derive(Clone, Copy)]
pub(in crate::doc) struct NativeFrameFacts {
    frame: Frame,
    origins: FrameOrigins,
}

impl NativeFrameFacts {
    pub(super) fn new(frame: Frame, origins: FrameOrigins) -> Self {
        Self { frame, origins }
    }

    #[allow(
        dead_code,
        reason = "native provenance seam; canonical frame mapping is pending"
    )]
    pub(in crate::doc) fn origin(&self, property: FrameProperty) -> Option<FrameOrigin> {
        self.origins[property as usize]
    }

    /// Native diagnostic value, not a shared pixel/frame equality predicate.
    #[allow(
        dead_code,
        reason = "native provenance seam; canonical frame mapping is pending"
    )]
    pub(in crate::doc) fn value(&self, property: FrameProperty) -> Option<i32> {
        Some(match property {
            FrameProperty::X => i32::from(self.frame.dxa_abs),
            FrameProperty::Y => i32::from(self.frame.dya_abs),
            FrameProperty::Width => i32::from(self.frame.width),
            FrameProperty::Height => i32::from(self.frame.height),
            FrameProperty::Anchors => i32::from(self.frame.position_code?),
            FrameProperty::Wrap => i32::from(self.frame.wrap),
            FrameProperty::HorizontalSpace => i32::from(self.frame.dxa_from_text),
            FrameProperty::VerticalSpace => i32::from(self.frame.dya_from_text),
            FrameProperty::Locked => i32::from(self.frame.locked),
            FrameProperty::NoOverlap => i32::from(self.frame.no_allow_overlap),
            FrameProperty::DropCap => i32::from(self.frame.drop_cap),
            FrameProperty::TextFlow => i32::from(self.frame.text_flow?),
        })
    }

    pub(in crate::doc) fn table_paragraph_facts(&self) -> Option<TableParagraphFrame> {
        self.frame.table_paragraph_facts()
    }

    /// [MS-DOC] 2.4.3 row-identity key; `None` facts are the unframed default.
    pub(in crate::doc) fn table_row_key(
        facts: Option<&Self>,
    ) -> crate::doc::table_structure::FrameKey {
        facts.map_or_else(
            || Frame::default().table_row_key(),
            |facts| facts.frame.table_row_key(),
        )
    }
}
