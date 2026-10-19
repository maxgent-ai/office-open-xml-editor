//! Incremental work accounting for legacy embedded OPC entries. MS-ODRAW
//! §2.3.4.42 and MS-XLS §2.4.326 embed ZIP packages; this is library resource
//! policy, not a compressed-size validity rule or admission check.
use std::io::Read;

fn unreadable() -> String {
    crate::rejection::unsupported("invalid legacy embedded ZIP entry data")
}
fn exhausted() -> String {
    crate::rejection::resource("legacy embedded ZIP input work budget exceeded")
}

/// Read the raw entry so a decoder call cannot hide arbitrarily much
/// zero-output work. Each call receives at most the remaining input allowance;
/// total_in charges only bytes actually consumed, even on a decoder error.
/// Stored entries spend the same work per consumed byte. Skipped archive
/// prefixes, unread entries and buffered but unconsumed input cost no work.
/// Output/retention budgets remain the caller's responsibility. CRC and size
/// integrity are checked at EOF, after actual input/output accounting.
pub(crate) struct Entry<'a, R: Read> {
    raw: zip::read::ZipFile<'a, R>,
    decoder: Option<flate2::Decompress>,
    buffer: [u8; 8192],
    start: usize,
    end: usize,
    crc: flate2::Crc,
    actual: u64,
    finished: bool,
}

impl<'a, R: Read + std::io::Seek> Entry<'a, R> {
    pub(crate) fn open(archive: &'a mut zip::ZipArchive<R>, index: usize) -> Result<Self, String> {
        let raw = archive.by_index_raw(index).map_err(|_| unreadable())?;
        if raw.encrypted() {
            return Err(crate::rejection::encryption(
                "encrypted legacy embedded ZIP entry",
            ));
        }
        let decoder = match raw.compression() {
            zip::CompressionMethod::Stored => None,
            zip::CompressionMethod::Deflated => Some(flate2::Decompress::new(false)),
            _ => return Err(unreadable()),
        };
        Ok(Self {
            raw,
            decoder,
            buffer: [0; 8192],
            start: 0,
            end: 0,
            crc: flate2::Crc::new(),
            actual: 0,
            finished: false,
        })
    }

    pub(crate) fn size(&self) -> u64 {
        self.raw.size()
    }

    pub(crate) fn read(&mut self, output: &mut [u8], work: &mut usize) -> Result<usize, String> {
        if output.is_empty() {
            return Ok(0);
        }
        if self.finished {
            return self.finish();
        }
        loop {
            if self.start == self.end {
                // Bounded read-ahead is not decoder consumption. In particular,
                // an EOF probe must allow an empty Stored entry at zero work.
                self.end = self.raw.read(&mut self.buffer).map_err(|_| unreadable())?;
                self.start = 0;
            }
            let available = self.end - self.start;
            let allowance = available.min(*work);
            let (consumed, produced, finished, valid) = if let Some(decoder) = &mut self.decoder {
                let before_in = decoder.total_in();
                let before_out = decoder.total_out();
                let status = decoder.decompress(
                    &self.buffer[self.start..self.start + allowance],
                    output,
                    if available == 0 {
                        flate2::FlushDecompress::Finish
                    } else {
                        flate2::FlushDecompress::None
                    },
                );
                (
                    (decoder.total_in() - before_in) as usize,
                    (decoder.total_out() - before_out) as usize,
                    matches!(status, Ok(flate2::Status::StreamEnd)),
                    status.is_ok(),
                )
            } else {
                let count = allowance.min(output.len());
                output[..count].copy_from_slice(&self.buffer[self.start..self.start + count]);
                (count, count, available == 0, true)
            };
            *work -= consumed; // decoder was offered no more than this allowance
            self.start += consumed;
            self.actual += produced as u64;
            self.crc.update(&output[..produced]);
            if !valid {
                return Err(unreadable());
            }
            if finished {
                self.finished = true;
            }
            if produced != 0 {
                return Ok(produced);
            }
            if finished {
                return self.finish();
            }
            if consumed == 0 {
                // Pending output/end markers can be resolved with no new input.
                // Only reject once progress needs input beyond the allowance.
                return Err(if available != 0 && *work == 0 {
                    exhausted()
                } else {
                    unreadable()
                });
            }
        }
    }

    fn finish(&self) -> Result<usize, String> {
        if self.actual != self.raw.size() || self.crc.sum() != self.raw.crc32() {
            return Err(unreadable());
        }
        Ok(0)
    }

    pub(crate) fn read_to_end(&mut self, work: &mut usize, limit: u64) -> Result<Vec<u8>, String> {
        let mut bytes = Vec::new();
        let mut chunk = [0; 8192];
        loop {
            let allowance = (limit - bytes.len() as u64)
                .saturating_add(1)
                .min(chunk.len() as u64) as usize;
            let count = self.read(&mut chunk[..allowance], work)?;
            if count == 0 {
                return Ok(bytes);
            }
            if bytes.len() as u64 + count as u64 > limit {
                return Err(crate::rejection::resource(
                    "legacy embedded ZIP part byte budget exceeded",
                ));
            }
            bytes.try_reserve(count).map_err(|_| {
                crate::rejection::resource("legacy embedded ZIP part allocation failed")
            })?;
            bytes.extend_from_slice(&chunk[..count]);
        }
    }

    #[cfg(feature = "direct-ppt")]
    pub(crate) fn read_to_string(
        &mut self,
        work: &mut usize,
        limit: u64,
    ) -> Result<String, String> {
        String::from_utf8(self.read_to_end(work, limit)?).map_err(|_| unreadable())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn tiny_output_spends_input_work_and_accepts_the_exact_boundary() {
        let bytes = crate::opc_part::test_empty_deflate("part.xml", 2000, b"x");
        let mut zip = zip::ZipArchive::new(std::io::Cursor::new(&bytes)).unwrap();
        let mut work = 100;
        let error = Entry::open(&mut zip, 0)
            .unwrap()
            .read_to_end(&mut work, 1)
            .unwrap_err();
        assert!(error.starts_with("LEGACY_RESOURCE_LIMIT:"), "{error}");
        assert_eq!(work, 0);
        let mut work = 10_001; // 2000 stored-block headers and one payload byte
        assert_eq!(
            Entry::open(&mut zip, 0)
                .unwrap()
                .read_to_end(&mut work, 1)
                .unwrap(),
            b"x"
        );
        assert_eq!(work, 0);
    }

    #[test]
    fn raw_entry_decoding_preserves_crc_and_truncation_rejections() {
        let valid = crate::opc_part::test_empty_deflate("part.xml", 1, b"x");
        let data = 30 + "part.xml".len();
        for corrupt_payload in [false, true] {
            let mut bytes = valid.clone();
            if corrupt_payload {
                bytes[data + 5] = b'y'; // valid DEFLATE, wrong CRC
            } else {
                bytes[data] = 0; // non-final block with no following block
            }
            let mut zip = zip::ZipArchive::new(std::io::Cursor::new(&bytes)).unwrap();
            let error = Entry::open(&mut zip, 0)
                .unwrap()
                .read_to_end(&mut 100, 1)
                .unwrap_err();
            assert!(error.starts_with("UNSUPPORTED:"), "{error}");
        }
    }
}
