//! Display size of a `<w:numPicBullet>` picture bullet (ECMA-376 §17.9.20).
//!
//! §17.9.20 says only that the VML picture is the bullet; it defines no
//! display size. Observed Windows Word (16.0) behavior, measured on Word's
//! PDFs of authored DOCX controls (VML `<v:shape style>` extent, raster of a
//! known pixel size and density, marker run size from the level's `w:rPr`
//! over the paragraph mark):
//! - the bullet is sized automatically from the marker run's size, like the
//!   DOC automatic picture bullet: its height is the marker size minus 2 pt,
//!   never less than 2 pt (markers 1, 2, 3, 4, 5, 8, 42, 72 and 144 pt);
//! - that height is then scaled by the VML height over the picture's natural
//!   height, and the width follows the VML aspect ratio. A 9 pt VML extent on
//!   a 32 px image without density (24 pt natural) shows 15 pt at a 42 pt
//!   marker and 2.25 pt at an 8 pt one; VML 18 pt doubles it; VML 9x18 pt
//!   shows 15x30 pt; a 64 px image halves it; 192 dpi doubles it; a 48x24 px
//!   image with a 9x9 pt VML extent shows 20x20 pt;
//! - the natural size is pixels at the image's own density: PNG pHYs, JPEG
//!   JFIF density, otherwise 96 dpi (PNG without pHYs, JPEG with an aspect-only
//!   JFIF, GIF);
//! - a VML dimension that the style omits acts as 50 pt (a missing height on
//!   24, 48 and 12 pt natural images, and a missing height beside a 9 pt
//!   width, all fit that value);
//! - a complex-script or right-to-left marker (`w:cs`, `w:rtl`) uses its
//!   `w:szCs` size.
//!
//! Each displayed dimension is limited to 32767 twips, as for DOC picture
//! bullets. Other image formats (BMP, TIFF, metafiles) have no measured
//! natural size here, so their bullets keep the VML extent as authored.

/// 32767 twips.
const MAX_DIMENSION_PT: f64 = 1638.35;
/// A VML picture-bullet dimension the style omits.
const ABSENT_VML_DIMENSION_PT: f64 = 50.0;
const DEFAULT_DPI: f64 = 96.0;

/// The display box of an automatic DOCX picture bullet, or `None` when the
/// picture has no measured natural size.
pub(crate) fn display_box(
    marker_size_pt: f64,
    vml_width_pt: Option<f64>,
    vml_height_pt: Option<f64>,
    natural_pt: (f64, f64),
) -> Option<(f64, f64)> {
    let (_, natural_height) = natural_pt;
    if !(natural_height.is_finite() && natural_height > 0.0 && marker_size_pt.is_finite()) {
        return None;
    }
    let vml_width = vml_width_pt.unwrap_or(ABSENT_VML_DIMENSION_PT);
    let vml_height = vml_height_pt.unwrap_or(ABSENT_VML_DIMENSION_PT);
    if !(vml_width > 0.0 && vml_height > 0.0) {
        return None;
    }
    let height = (marker_size_pt - 2.0).max(2.0) * vml_height / natural_height;
    let width = height * vml_width / vml_height;
    Some((width.min(MAX_DIMENSION_PT), height.min(MAX_DIMENSION_PT)))
}

/// Natural size in points of a PNG, JPEG or GIF picture.
pub(crate) fn natural_size_pt(bytes: &[u8]) -> Option<(f64, f64)> {
    let (width, height, dpi_x, dpi_y) = if bytes.starts_with(b"\x89PNG\r\n\x1a\n") {
        png(bytes)?
    } else if bytes.starts_with(&[0xff, 0xd8]) {
        jpeg(bytes)?
    } else if bytes.starts_with(b"GIF87a") || bytes.starts_with(b"GIF89a") {
        let width = u16::from_le_bytes([*bytes.get(6)?, *bytes.get(7)?]);
        let height = u16::from_le_bytes([*bytes.get(8)?, *bytes.get(9)?]);
        (
            u32::from(width),
            u32::from(height),
            DEFAULT_DPI,
            DEFAULT_DPI,
        )
    } else {
        return None;
    };
    if width == 0 || height == 0 {
        return None;
    }
    Some((
        f64::from(width) * 72.0 / dpi_x,
        f64::from(height) * 72.0 / dpi_y,
    ))
}

fn be32(bytes: &[u8], at: usize) -> Option<u32> {
    Some(u32::from_be_bytes(bytes.get(at..at + 4)?.try_into().ok()?))
}

fn be16(bytes: &[u8], at: usize) -> Option<u16> {
    Some(u16::from_be_bytes(bytes.get(at..at + 2)?.try_into().ok()?))
}

/// PNG IHDR size and pHYs density (unit 1: pixels per metre).
fn png(bytes: &[u8]) -> Option<(u32, u32, f64, f64)> {
    let mut at = 8;
    let mut size = None;
    let mut dpi = (DEFAULT_DPI, DEFAULT_DPI);
    while at + 8 <= bytes.len() {
        let length = be32(bytes, at)? as usize;
        let kind = bytes.get(at + 4..at + 8)?;
        let data = bytes.get(at + 8..(at + 8).checked_add(length)?)?;
        match kind {
            b"IHDR" if data.len() >= 8 => size = Some((be32(data, 0)?, be32(data, 4)?)),
            b"pHYs" if data.len() >= 9 && data[8] == 1 => {
                let (x, y) = (be32(data, 0)?, be32(data, 4)?);
                if x > 0 && y > 0 {
                    dpi = (f64::from(x) * 0.0254, f64::from(y) * 0.0254);
                }
            }
            b"IDAT" | b"IEND" => break,
            _ => {}
        }
        at = at.checked_add(12)?.checked_add(length)?;
    }
    let (width, height) = size?;
    Some((width, height, dpi.0, dpi.1))
}

/// JPEG frame size and JFIF density (units 1: dpi, 2: dots per cm).
fn jpeg(bytes: &[u8]) -> Option<(u32, u32, f64, f64)> {
    let mut at = 2;
    let mut dpi = (DEFAULT_DPI, DEFAULT_DPI);
    while at + 4 <= bytes.len() {
        if bytes[at] != 0xff {
            return None;
        }
        let marker = bytes[at + 1];
        if marker == 0xff {
            at += 1;
            continue;
        }
        let length = usize::from(be16(bytes, at + 2)?);
        let segment = bytes.get(at + 4..at + 2 + length)?;
        match marker {
            0xe0 if segment.starts_with(b"JFIF\0") && segment.len() >= 12 => {
                let units = segment[7];
                let (x, y) = (be16(segment, 8)?, be16(segment, 10)?);
                if x > 0 && y > 0 && matches!(units, 1 | 2) {
                    let scale = if units == 2 { 2.54 } else { 1.0 };
                    dpi = (f64::from(x) * scale, f64::from(y) * scale);
                }
            }
            0xc0..=0xcf if !matches!(marker, 0xc4 | 0xc8 | 0xcc) && segment.len() >= 5 => {
                let height = u32::from(be16(segment, 1)?);
                let width = u32::from(be16(segment, 3)?);
                return Some((width, height, dpi.0, dpi.1));
            }
            0xd9 | 0xda => return None,
            _ => {}
        }
        at += 2 + length;
    }
    None
}

#[cfg(test)]
mod tests {
    use super::*;

    fn png_with(width: u32, height: u32, phys_ppm: Option<u32>) -> Vec<u8> {
        let mut bytes = b"\x89PNG\r\n\x1a\n".to_vec();
        let mut chunk = |kind: &[u8], data: &[u8]| {
            bytes.extend((data.len() as u32).to_be_bytes());
            bytes.extend(kind);
            bytes.extend(data);
            bytes.extend([0; 4]);
        };
        let mut ihdr = width.to_be_bytes().to_vec();
        ihdr.extend(height.to_be_bytes());
        ihdr.extend([8, 2, 0, 0, 0]);
        chunk(b"IHDR", &ihdr);
        if let Some(ppm) = phys_ppm {
            let mut phys = ppm.to_be_bytes().to_vec();
            phys.extend(ppm.to_be_bytes());
            phys.push(1);
            chunk(b"pHYs", &phys);
        }
        chunk(b"IEND", &[]);
        bytes
    }

    #[test]
    fn natural_size_follows_the_image_density() {
        assert_eq!(natural_size_pt(&png_with(32, 32, None)), Some((24.0, 24.0)));
        // 7559 px/m is 192 dpi.
        let (w, h) = natural_size_pt(&png_with(32, 32, Some(7559))).unwrap();
        assert!((w - 12.0).abs() < 0.01 && (h - 12.0).abs() < 0.01);
        let gif = [b"GIF89a".as_slice(), &[64, 0, 32, 0]].concat();
        assert_eq!(natural_size_pt(&gif), Some((48.0, 24.0)));
        let jfif = |units: u8, density: u16| {
            let mut bytes = vec![0xff, 0xd8, 0xff, 0xe0, 0, 16];
            bytes.extend(b"JFIF\0\x01\x01");
            bytes.push(units);
            bytes.extend(density.to_be_bytes());
            bytes.extend(density.to_be_bytes());
            bytes.extend([0, 0]);
            bytes.extend([0xff, 0xc0, 0, 11, 8, 0, 32, 0, 32, 3, 0, 0, 0, 0]);
            bytes
        };
        let (w, _) = natural_size_pt(&jfif(1, 150)).unwrap();
        assert!((w - 15.36).abs() < 0.001);
        assert_eq!(natural_size_pt(&jfif(0, 1)), Some((24.0, 24.0)));
        assert_eq!(natural_size_pt(b"BM...."), None);
        // A chunk length past the end of the data (or of usize) is rejected.
        let mut huge = b"\x89PNG\r\n\x1a\n".to_vec();
        huge.extend(0xffff_fff0u32.to_be_bytes());
        huge.extend(b"tEXt");
        huge.extend([0; 8]);
        assert_eq!(natural_size_pt(&huge), None);
    }

    #[test]
    fn display_box_matches_the_word_controls() {
        let natural = (24.0, 24.0);
        let close = |a: Option<(f64, f64)>, b: (f64, f64)| {
            let (x, y) = a.unwrap();
            assert!(
                (x - b.0).abs() < 0.01 && (y - b.1).abs() < 0.01,
                "{x}x{y} vs {b:?}"
            );
        };
        close(
            display_box(42.0, Some(9.0), Some(9.0), natural),
            (15.0, 15.0),
        );
        close(
            display_box(8.0, Some(9.0), Some(9.0), natural),
            (2.25, 2.25),
        );
        close(
            display_box(42.0, Some(18.0), Some(18.0), natural),
            (30.0, 30.0),
        );
        close(
            display_box(42.0, Some(9.0), Some(18.0), natural),
            (15.0, 30.0),
        );
        close(
            display_box(42.0, Some(9.0), Some(9.0), (48.0, 48.0)),
            (7.5, 7.5),
        );
        close(
            display_box(42.0, Some(9.0), Some(9.0), (36.0, 18.0)),
            (20.0, 20.0),
        );
        // The 2 pt floor applies before scaling.
        close(
            display_box(1.0, Some(9.0), Some(9.0), natural),
            (0.75, 0.75),
        );
        // Absent VML dimensions act as 50 pt.
        close(display_box(42.0, None, None, natural), (83.33, 83.33));
        close(display_box(42.0, Some(9.0), None, natural), (15.0, 83.33));
        close(
            display_box(1638.0, Some(9.0), Some(9.0), (1.0, 1.0)),
            (1638.35, 1638.35),
        );
    }
}
