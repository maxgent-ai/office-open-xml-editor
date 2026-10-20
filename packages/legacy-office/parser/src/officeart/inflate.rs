//! Exact zlib inflation for OfficeArt metafile BLIPs (MS-ODRAW 2.2.31),
//! charged by produced bytes. Dependency-free apart from flate2 so the
//! allocation regression test (`tests/metafile_inflate_alloc.rs`) can include
//! it under a counting global allocator.

/// Output growth step. Memory follows the bytes the stream actually
/// produces, never the declared cbSize or the compressed cbSave.
pub(crate) const INFLATE_CHUNK: usize = 64 * 1024;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum InflateError {
    /// The DEFLATE stream (or its RFC 1950 wrapper/checksum) is corrupt.
    Invalid,
    /// The stream did not produce exactly `expanded` bytes from exactly
    /// `stored` input bytes: a structural mismatch with the header.
    SizeMismatch,
    /// The stream produced more than `cap` bytes.
    OverBudget,
}

/// Inflate one RFC 1950 zlib-wrapped DEFLATE stream that must produce exactly
/// `expanded` bytes from exactly `stored` input bytes, producing at most
/// `cap` bytes.
///
/// The destination grows by [`INFLATE_CHUNK`] as output is produced, so a
/// header that declares a large size over a small stream never reserves the
/// declared size (allocating it up front would permanently grow WASM linear
/// memory). Production stops one byte past `min(expanded, cap)`: that spare
/// byte detects a stream longer than its declaration, or one that exceeds the
/// cap, without letting it grow further. The declared sizes are only
/// compared with what the stream actually does; neither admits nor rejects
/// input by itself.
pub(crate) fn inflate_exact(
    bytes: &[u8],
    expanded: usize,
    stored: usize,
    cap: usize,
) -> Result<Vec<u8>, InflateError> {
    let bound = expanded.min(cap);
    let limit = bound.saturating_add(1);
    let mut output = Vec::new();
    let mut decoder = flate2::Decompress::new(true);
    let mut status = flate2::Status::Ok;
    while output.len() < limit && status != flate2::Status::StreamEnd {
        let produced_before = output.len();
        let consumed_before = decoder.total_in() as usize;
        let room = (limit - produced_before).min(INFLATE_CHUNK);
        if output.capacity() < produced_before + room {
            // Amortized doubling, but never beyond the production limit.
            let target = (output.capacity() * 2)
                .max(produced_before + room)
                .min(limit);
            output.reserve_exact(target - produced_before);
        }
        output.resize(produced_before + room, 0);
        status = decoder
            .decompress(
                &bytes[consumed_before..],
                &mut output[produced_before..],
                flate2::FlushDecompress::None,
            )
            .map_err(|_| InflateError::Invalid)?;
        // `total_out` counts every byte written so far, i.e. the kept prefix.
        output.truncate(decoder.total_out() as usize);
        if output.len() == produced_before && decoder.total_in() as usize == consumed_before {
            break; // Truncated stream: no further progress is possible.
        }
    }
    if output.len() > cap {
        return Err(InflateError::OverBudget);
    }
    if status != flate2::Status::StreamEnd
        || decoder.total_out() != expanded as u64
        || decoder.total_in() != stored as u64
    {
        return Err(InflateError::SizeMismatch);
    }
    output.shrink_to_fit();
    Ok(output)
}
