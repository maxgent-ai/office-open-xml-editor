//! The Bullet Pictures document: [MS-DOC] 2.6.1 sprmCPbiIBullet locates a
//! picture bullet by CP relative to the hidden bookmark `_PictureBullets`
//! in the main document. Bookmarks are read from SttbfBkmk (2.9.279) with
//! the parallel Plcfbkf (2.8.10, FBKF 2.9.70) and Plcfbkl (2.8.12).
//!
//! Acquisition is lazy-failing: a malformed or missing bookmark is recorded,
//! not raised, so documents whose bullets are never enabled keep their
//! existing result. Only an enabled picture bullet resolves the document.
//! Bookmarks, fields and names are data; nothing here executes them.

use super::{fkp, u16_at, u32_at, unsupported, Story};

/// FibRgFcLcb97 fcSttbfBkmk, fcPlcfBkf and fcPlcfBkl (entries 21-23 after
/// the FibRgFcLcb97 start at 0x9A).
const STTBF_BKMK: usize = 0x142;
const PLCF_BKF: usize = 0x14a;
const PLCF_BKL: usize = 0x152;
/// [MS-DOC] 2.9.279: at most 0x3FFB bookmark names.
const MAX_BOOKMARKS: usize = 0x3ffb;
/// Library resource policy: the Bullet Pictures document holds one picture
/// character per bullet; larger spans are not retained.
const MAX_BULLET_DOCUMENT_UNITS: usize = 65_536;
const NAME: &str = "_PictureBullets";

/// One character of the Bullet Pictures document with its formatting
/// location: the piece FC/PRM and the FC of its paragraph's mark.
#[derive(Clone, Copy, Debug)]
pub(super) struct Character<'a> {
    pub(super) unit: u16,
    pub(super) fc: usize,
    pub(super) prm: u16,
    pub(super) prc: Option<&'a [u8]>,
    pub(super) mark_fc: usize,
}

#[derive(Debug)]
pub(super) struct Document<'a> {
    characters: Result<Vec<Character<'a>>, String>,
}

impl<'a> Document<'a> {
    /// `None` when the document has no `_PictureBullets` bookmark.
    pub(super) fn read(word: &[u8], table: &[u8], story: &Story<'a>) -> Option<Self> {
        match find(word, table) {
            Ok(None) => None,
            Ok(Some(range)) => Some(Self {
                characters: characters(story, range),
            }),
            Err(error) => Some(Self {
                characters: Err(error),
            }),
        }
    }

    /// The character at `relative_cp` from the bookmark start. A position at
    /// or past the bookmark end is outside the Bullet Pictures document.
    pub(super) fn character(&self, relative_cp: u32) -> Result<Character<'a>, String> {
        let characters = self.characters.as_ref().map_err(Clone::clone)?;
        usize::try_from(relative_cp)
            .ok()
            .and_then(|index| characters.get(index))
            .copied()
            .ok_or_else(|| unsupported("Word picture bullet outside its bookmark"))
    }
}

/// The CP range of the `_PictureBullets` bookmark, if present.
fn find(word: &[u8], table: &[u8]) -> Result<Option<std::ops::Range<usize>>, String> {
    let names = fkp::table_part(word, table, STTBF_BKMK)?;
    if names.is_empty() {
        return Ok(None);
    }
    let starts = fkp::table_part(word, table, PLCF_BKF)?;
    let limits = fkp::table_part(word, table, PLCF_BKL)?;
    // SttbfBkmk: extended STTB (fExtend 0xFFFF), 2-byte cData, cbExtra 0.
    if names.len() < 6 || u16_at(names, 0)? != 0xffff || u16_at(names, 4)? != 0 {
        return Err(unsupported("invalid Word bookmark name table"));
    }
    let count = usize::from(u16_at(names, 2)?);
    if count > MAX_BOOKMARKS {
        return Err(unsupported("Word bookmark count budget exceeded"));
    }
    // Plcfbkf: (count + 1) CPs and count 4-byte FBKFs; Plcfbkl: count + 1 CPs.
    if starts.len() != (count + 1) * 4 + count * 4 || limits.len() != (count + 1) * 4 {
        return Err(unsupported(
            "Word bookmark tables disagree with their names",
        ));
    }
    let mut offset = 6usize;
    let mut found = None;
    for index in 0..count {
        let length = usize::from(u16_at(names, offset)?);
        let end = offset + 2 + length * 2;
        let units = names
            .get(offset + 2..end)
            .ok_or_else(|| unsupported("truncated Word bookmark name"))?;
        // Names MUST be 1..40 characters and unique (2.9.279).
        if !(1..40).contains(&length) {
            return Err(unsupported("invalid Word bookmark name length"));
        }
        offset = end;
        let matches = units
            .chunks_exact(2)
            .map(|unit| u16::from_le_bytes([unit[0], unit[1]]))
            .eq(NAME.encode_utf16());
        if !matches {
            continue;
        }
        if found.is_some() {
            return Err(unsupported("duplicate Word picture bullet bookmark"));
        }
        let start = u32_at(starts, index * 4)? as usize;
        let ibkl = usize::from(u16_at(starts, (count + 1) * 4 + index * 4)?);
        if ibkl >= count {
            return Err(unsupported("Word bookmark end outside its table"));
        }
        let limit = u32_at(limits, ibkl * 4)? as usize;
        if start > limit {
            return Err(unsupported("Word bookmark ends before it starts"));
        }
        found = Some(start..limit);
    }
    if offset != names.len() {
        return Err(unsupported("trailing Word bookmark name data"));
    }
    Ok(found)
}

fn characters<'a>(
    story: &Story<'a>,
    range: std::ops::Range<usize>,
) -> Result<Vec<Character<'a>>, String> {
    if range.len() > MAX_BULLET_DOCUMENT_UNITS {
        return Err(unsupported("Word picture bullet document budget exceeded"));
    }
    // The bookmark lies in the main document; its paragraph marks may follow
    // its end. One forward pass finds each character's paragraph mark.
    let units: Vec<u16> = story
        .text
        .encode_utf16()
        .skip(range.start)
        .take(MAX_BULLET_DOCUMENT_UNITS + 1)
        .take_while({
            let mut ended = false;
            let mut cp = range.start;
            move |unit| {
                let take = !ended;
                if cp >= range.end && matches!(unit, 0x0d | 0x07 | 0x0c) {
                    ended = true;
                }
                cp += 1;
                take
            }
        })
        .collect();
    if units.len() < range.len() {
        return Err(unsupported(
            "Word picture bullet bookmark outside the main story",
        ));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(range.len())
        .map_err(|_| "OUTPUT_TOO_LARGE".to_string())?;
    let mut mark_fc = None;
    for (index, unit) in units.iter().enumerate().rev() {
        let cp = range.start + index;
        let (_, fc, piece) = story
            .position(cp)
            .ok_or_else(|| unsupported("Word picture bullet outside the piece table"))?;
        if matches!(unit, 0x0d | 0x07 | 0x0c) {
            mark_fc = Some(fc);
        }
        if cp < range.end {
            let prc = if piece.prm & 1 != 0 {
                Some(
                    *story
                        .prcs
                        .get(usize::from(piece.prm >> 1))
                        .ok_or_else(|| unsupported("Word piece property index outside CLX"))?,
                )
            } else {
                None
            };
            output.push(Character {
                unit: *unit,
                fc,
                prm: piece.prm,
                prc,
                mark_fc: mark_fc
                    .ok_or_else(|| unsupported("Word picture bullet paragraph lacks a mark"))?,
            });
        }
    }
    output.reverse();
    Ok(output)
}
