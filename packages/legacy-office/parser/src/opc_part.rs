//! OPC part lookup for ZIPs embedded in legacy binary records.
//!
//! ECMA-376 Part 2 §6.2.2.3 gives part names ASCII case and RFC 3986
//! percent-encoding equivalence. Reject an ambiguous embedded ZIP rather than
//! choosing whichever duplicate its index happens to return.

pub(crate) fn entry_index<R: std::io::Read + std::io::Seek>(
    archive: &zip::ZipArchive<R>,
    path: &str,
) -> Option<usize> {
    let key = ooxml_common::rels::part_name_equivalence_key(path);
    let mut found = None;
    // `file_names` enumerates items in index order. A name the ZIP crate
    // cannot decode never equals a requested OPC part name.
    for (index, name) in archive.file_names().enumerate() {
        let Ok(name) = name else { continue };
        if ooxml_common::rels::part_name_equivalence_key(&name) == key {
            if found.is_some() {
                return None;
            }
            found = Some(index);
        }
    }
    found
}

mod reader;
pub(crate) use reader::Entry;

/// A valid ZIP with many empty DEFLATE stored blocks before its final output.
#[cfg(test)]
pub(crate) fn test_empty_deflate(name: &str, blocks: usize, output: &[u8]) -> Vec<u8> {
    use std::io::Write;
    let mut writer = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
    writer
        .start_file(
            name,
            zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Deflated),
        )
        .unwrap();
    writer.write_all(output).unwrap();
    let mut bytes = writer.finish().unwrap().into_inner();
    let central = bytes.windows(4).position(|s| s == b"PK\x01\x02").unwrap();
    let data = 30
        + u16::from_le_bytes(bytes[26..28].try_into().unwrap()) as usize
        + u16::from_le_bytes(bytes[28..30].try_into().unwrap()) as usize;
    let length = u16::try_from(output.len()).unwrap();
    let mut encoded = [0, 0, 0, 255, 255].repeat(blocks);
    let last = (blocks - 1) * 5;
    encoded[last] = 1; // BFINAL; earlier blocks produce no output.
    encoded[last + 1..last + 3].copy_from_slice(&length.to_le_bytes());
    encoded[last + 3..last + 5].copy_from_slice(&(!length).to_le_bytes());
    encoded.extend_from_slice(output);
    bytes.splice(data..central, encoded.iter().copied());
    let central = data + encoded.len();
    bytes[18..22].copy_from_slice(&(encoded.len() as u32).to_le_bytes());
    bytes[central + 20..central + 24].copy_from_slice(&(encoded.len() as u32).to_le_bytes());
    let footer = bytes.len() - 22;
    bytes[footer + 16..footer + 20].copy_from_slice(&(central as u32).to_le_bytes());
    bytes
}

/// Test fixture: splice `extra` into the first central-directory entry of
/// `zip` (central directory and EOCD sizes adjusted; offsets are unchanged
/// because the central directory follows every local entry).
#[cfg(test)]
pub(crate) fn test_with_central_extra(zip: &[u8], extra: &[u8]) -> Vec<u8> {
    let central = zip
        .windows(4)
        .position(|w| w == b"PK\x01\x02")
        .expect("central header");
    let eocd = zip
        .windows(4)
        .rposition(|w| w == b"PK\x05\x06")
        .expect("EOCD");
    let name_len = u16::from_le_bytes([zip[central + 28], zip[central + 29]]) as usize;
    let extra_len = u16::from_le_bytes([zip[central + 30], zip[central + 31]]) as usize;
    let at = central + 46 + name_len + extra_len;
    let mut out = zip[..at].to_vec();
    out.extend_from_slice(extra);
    out.extend_from_slice(&zip[at..]);
    let new_extra = (extra_len + extra.len()) as u16;
    out[central + 30..central + 32].copy_from_slice(&new_extra.to_le_bytes());
    let eocd = eocd + extra.len();
    let size =
        u32::from_le_bytes(out[eocd + 12..eocd + 16].try_into().unwrap()) + extra.len() as u32;
    out[eocd + 12..eocd + 16].copy_from_slice(&size.to_le_bytes());
    out
}

/// Test fixture: one Info-ZIP Unicode Path field restating `name`.
#[cfg(test)]
pub(crate) fn test_unicode_path_field(name: &str) -> Vec<u8> {
    let mut body = vec![1u8];
    let mut crc = flate2::Crc::new();
    crc.update(name.as_bytes());
    body.extend_from_slice(&crc.sum().to_le_bytes());
    body.extend_from_slice(name.as_bytes());
    let mut field = 0x7075u16.to_le_bytes().to_vec();
    field.extend_from_slice(&(body.len() as u16).to_le_bytes());
    field.extend_from_slice(&body);
    field
}
