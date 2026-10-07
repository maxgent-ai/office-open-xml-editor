//! Native DOC COLORREF acquisition shared by text, shading and borders.
//! [MS-DOC] 2.9.43 requires fAuto to be 0x00 (RGB) or 0xFF (automatic).
//! This differs from OfficeArtCOLORREF's flag bits; do not use it for drawings.

use super::unsupported;

/// None is cvAuto, not an absent color. Unused RGB bytes of cvAuto are only
/// SHOULD-zero in MS-DOC, so retain its automatic meaning without rejecting
/// nonzero RGB. Operand-specific NilBrc/ShdNil sentinels must be handled by
/// their callers before interpreting a COLORREF.
pub(super) fn read(bytes: &[u8]) -> Result<Option<[u8; 3]>, String> {
    match bytes.get(3) {
        Some(0) => Ok(Some([bytes[0], bytes[1], bytes[2]])),
        Some(0xff) => Ok(None),
        Some(_) => Err(unsupported("invalid Word COLORREF")),
        None => Err(unsupported("short Word COLORREF")),
    }
}
