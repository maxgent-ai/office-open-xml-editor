//! Index-only preflight stays local because ordinary OPC opening validates
//! part-size/ZIP64 declarations before inflation. Alternative output accounting
//! must precede those part-integrity rejections. Reuse the shared hard index
//! limits, but interpret only footer/index metadata here.
use super::unsupported;

/// Reserve index work before the ZIP crate walks entries. MS-ODRAW 2.3.4.42
/// embeds an OPC container; these are implementation metadata/work budgets, not limits on compressed or declared inflated bytes.
/// The footer count is the actual number of index records requested, including
/// empty entries. The borrowed directory walk bounds actual metadata bytes.
/// Keeping this check allocation-free also rejects forged ZIP64 capacities
/// before the ZIP dependency can reserve its index from an untrusted count.
fn policy(blob: &[u8], work_budget: &mut usize) -> Result<Policy, String> {
    let unreadable = || unsupported("unreadable PowerPoint alternative shape XML package");
    let exhausted =
        || crate::rejection::resource("PowerPoint alternative ZIP metadata work budget exceeded");
    let u16_at = |at: usize| -> Option<u16> {
        Some(u16::from_le_bytes(
            blob.get(at..at.checked_add(2)?)?.try_into().ok()?,
        ))
    };
    let u32_at = |at: usize| -> Option<u32> {
        Some(u32::from_le_bytes(
            blob.get(at..at.checked_add(4)?)?.try_into().ok()?,
        ))
    };
    let u64_at = |at: usize| -> Option<u64> {
        Some(u64::from_le_bytes(
            blob.get(at..at.checked_add(8)?)?.try_into().ok()?,
        ))
    };
    // APPNOTE EOCD: 22 fixed bytes plus at most 65,535 comment bytes.
    let start = blob.len().saturating_sub(22 + u16::MAX as usize);
    let eocd = blob[start..]
        .windows(4)
        .enumerate()
        .rev()
        .find_map(|(relative, sig)| {
            let at = start + relative;
            (sig == b"PK\x05\x06"
                && at.checked_add(22 + u16_at(at + 20)? as usize) == Some(blob.len()))
            .then_some(at)
        })
        .ok_or_else(unreadable)?;
    let mut count = u16_at(eocd + 8).ok_or_else(unreadable)? as u64;
    let mut directory_end = eocd;
    let mut relative_offset = u32_at(eocd + 16).ok_or_else(unreadable)? as u64;
    let mut zip64_at = None;
    let mut directory_bytes = u32_at(eocd + 12).ok_or_else(unreadable)? as u64;
    // Match zip 2.x's ZIP64 switch exactly: a locator alone does not override
    // the classic per-disk count used for its index allocation.
    let uses_zip64 = u16_at(eocd + 10) == Some(u16::MAX) || relative_offset == u32::MAX as u64;
    if let Some(locator) = eocd
        .checked_sub(20)
        .filter(|at| uses_zip64 && blob.get(*at..*at + 4) == Some(b"PK\x06\x07"))
    {
        // The ZIP64 record may follow an inert SFX prefix and contain an
        // extensible-data sector. Unlike the fixed 65,557-byte EOCD window,
        // this backward search has no format bound. Charge each inspected
        // position before examining it; never charge skipped prefix bytes.
        let mut zip64 = None;
        for (at, sig) in blob[..locator].windows(4).enumerate().rev() {
            *work_budget = work_budget.checked_sub(1).ok_or_else(exhausted)?;
            if sig == b"PK\x06\x06"
                && u64_at(at + 4).is_some_and(|size| {
                    size >= 44
                        && (at as u64)
                            .checked_add(12)
                            .and_then(|end| end.checked_add(size))
                            == Some(locator as u64)
                })
            {
                zip64 = Some(at);
                break;
            }
        }
        let zip64 = zip64.ok_or_else(unreadable)?;
        // Charge the physically present footer extent before the ZIP crate
        // copies the extensible-data sector. A declaration beyond input is
        // malformed rather than a measured metadata resource crossing.
        if (eocd - zip64) as u64 > ooxml_common::resource::HARD_MAX_CENTRAL_DIRECTORY_BYTES {
            return Err(exhausted());
        }
        count = u64_at(zip64 + 32).ok_or_else(unreadable)?;
        directory_end = zip64;
        relative_offset = u64_at(zip64 + 48).ok_or_else(unreadable)?;
        zip64_at = Some(zip64);
        directory_bytes = u64_at(zip64 + 40).ok_or_else(unreadable)?;
    }
    let count = usize::try_from(count).map_err(|_| exhausted())?;
    if count as u64 > ooxml_common::resource::HARD_MAX_ARCHIVE_ENTRIES {
        return Err(exhausted());
    }
    *work_budget = work_budget.checked_sub(count).ok_or_else(exhausted)?;
    let directory_bytes = usize::try_from(directory_bytes).map_err(|_| unreadable())?;
    let start = directory_end
        .checked_sub(directory_bytes)
        .ok_or_else(unreadable)?;
    // This is an actual borrowed metadata extent, never an inflated-size gate.
    // Include retained filename duplication in this same metadata allowance.
    if directory_bytes as u64 > ooxml_common::resource::HARD_MAX_CENTRAL_DIRECTORY_BYTES {
        return Err(exhausted());
    }
    let mut cursor = start;
    let mut metadata_bytes = directory_bytes.saturating_add(blob.len() - directory_end);
    for _ in 0..count {
        if blob.get(cursor..cursor + 4) != Some(b"PK\x01\x02") {
            return Err(unreadable());
        }
        if u16_at(cursor + 8).ok_or_else(unreadable)? & 1 != 0 {
            return Err(crate::rejection::encryption(
                "encrypted PowerPoint alternative ZIP entry",
            ));
        }
        metadata_bytes =
            metadata_bytes.saturating_add(u16_at(cursor + 28).ok_or_else(unreadable)? as usize);
        if metadata_bytes as u64 > ooxml_common::resource::HARD_MAX_CENTRAL_DIRECTORY_BYTES {
            return Err(exhausted());
        }
        cursor = cursor
            .checked_add(
                46 + u16_at(cursor + 28).ok_or_else(unreadable)? as usize
                    + u16_at(cursor + 30).ok_or_else(unreadable)? as usize
                    + u16_at(cursor + 32).ok_or_else(unreadable)? as usize,
            )
            .ok_or_else(unreadable)?;
        if cursor > directory_end {
            return Err(unreadable());
        }
    }
    if metadata_bytes as u64 > ooxml_common::resource::HARD_MAX_CENTRAL_DIRECTORY_BYTES {
        return Err(exhausted());
    }
    if cursor != directory_end {
        // APPNOTE 4.3.13 permits a digital-signature record after the directory.
        if blob.get(cursor..cursor + 4) != Some(b"PK\x05\x05")
            || cursor.checked_add(6 + u16_at(cursor + 4).ok_or_else(unreadable)? as usize)
                != Some(directory_end)
        {
            return Err(unreadable());
        }
    }
    let archive_offset = (start as u64)
        .checked_sub(relative_offset)
        .ok_or_else(unreadable)?;
    let zip64_records: Vec<usize> = zip64_at
        .into_iter()
        .chain(eocd.checked_sub(20).filter(|_| zip64_at.is_some()))
        .collect();
    Ok(Policy {
        view: ooxml_common::zip::IndexView::for_validated_directory(
            blob,
            start,
            eocd,
            &zip64_records,
        ),
        archive_offset,
        entries: count,
    })
}

struct Policy {
    /// The validated footer and directory, as the ZIP reader may see them
    /// while indexing (shared `ooxml_common::zip::IndexView` rules).
    view: ooxml_common::zip::IndexView,
    archive_offset: u64,
    /// The number of directory records the preflight validated.
    entries: usize,
}

/// zip 2.x retries earlier footer signatures after an index error. Block that
/// route to an unbudgeted count or ZIP64 extensible-data allocation. A successful
/// open must use only the reserved footer. All subsequent legacy lookups use
/// this same construction path. The guard is disabled for decompression.
pub(super) struct IndexReader<'a> {
    cursor: std::io::Cursor<&'a [u8]>,
    policy: Policy,
    state: std::rc::Rc<std::cell::Cell<IndexState>>,
}

impl std::io::Read for IndexReader<'_> {
    fn read(&mut self, output: &mut [u8]) -> std::io::Result<usize> {
        if self.state.get() == IndexState::Ready {
            return std::io::Read::read(&mut self.cursor, output);
        }
        let at = self.cursor.position() as usize;
        let count = self
            .policy
            .view
            .read(self.cursor.get_ref(), at, output)
            .inspect_err(|_| self.state.set(IndexState::DifferentFooter))?;
        self.cursor
            .set_position((at.min(self.cursor.get_ref().len()) + count) as u64);
        Ok(count)
    }
}

impl std::io::Seek for IndexReader<'_> {
    fn seek(&mut self, position: std::io::SeekFrom) -> std::io::Result<u64> {
        std::io::Seek::seek(&mut self.cursor, position)
    }
}

#[derive(Clone, Copy, PartialEq)]
enum IndexState {
    Indexing,
    DifferentFooter,
    Ready,
}

/// Adoption supplies the session work ledger. Later bounded projection reads
/// reserve an independent index allowance, without recharging inflated output.
pub(super) fn open<'a>(
    blob: &'a [u8],
    work_budget: &mut usize,
) -> Result<zip::ZipArchive<IndexReader<'a>>, String> {
    let policy = policy(blob, work_budget)?;
    let expected_entries = policy.entries;
    let config = zip::read::Config {
        archive_offset: zip::read::ArchiveOffset::Known(policy.archive_offset),
    };
    let state = std::rc::Rc::new(std::cell::Cell::new(IndexState::Indexing));
    let reader = IndexReader {
        cursor: std::io::Cursor::new(blob),
        policy,
        state: state.clone(),
    };
    let unreadable = || unsupported("unreadable PowerPoint alternative shape XML package");
    let archive = zip::ZipArchive::with_config(config, reader).map_err(|_| unreadable())?;
    // The index must be exactly the validated directory.
    if state.get() == IndexState::DifferentFooter || archive.len() != expected_entries {
        return Err(unreadable());
    }
    state.set(IndexState::Ready);
    Ok(archive)
}

#[cfg(test)]
mod tests {
    #[test]
    fn repeated_unicode_path_fields_are_opaque_metadata() {
        use std::io::Write;
        let mut writer = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
        writer
            .start_file("a.xml", zip::write::SimpleFileOptions::default())
            .unwrap();
        writer.write_all(b"a").unwrap();
        let zip = writer.finish().unwrap().into_inner();
        let field = crate::opc_part::test_unicode_path_field("a.xml");
        let once = crate::opc_part::test_with_central_extra(&zip, &field);
        assert!(super::open(&once, &mut 1_000).is_ok());
        let twice =
            crate::opc_part::test_with_central_extra(&zip, &[field.clone(), field].concat());
        // Unicode metadata is opaque to the index (OPC names are raw ASCII).
        assert!(super::open(&twice, &mut 1_000).is_ok());
    }

    #[test]
    fn eocd_signature_bytes_in_the_comment_are_inert() {
        use std::io::Write;
        let mut writer = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
        writer
            .start_file("a.xml", zip::write::SimpleFileOptions::default())
            .unwrap();
        writer.write_all(b"a").unwrap();
        writer.set_comment("decoy PK\u{5}\u{6} record").unwrap();
        let zip = writer.finish().unwrap().into_inner();
        // Inert comment bytes are not a footer (OPC Annex B.2 allows comments).
        assert!(super::open(&zip, &mut 1_000).is_ok());
    }
}
