//! Complete native source acquisition: input length is not admission policy.
use std::io::{self, Read};

pub fn read_source(mut source: impl Read) -> io::Result<Vec<u8>> {
    let mut bytes = Vec::new();
    let mut chunk = [0; 64 * 1024];
    loop {
        let count = match source.read(&mut chunk) {
            Err(error) if error.kind() == io::ErrorKind::Interrupted => continue,
            result => result?,
        };
        if count == 0 {
            return Ok(bytes);
        }
        reserve(&mut bytes, count)?;
        bytes.extend_from_slice(&chunk[..count]);
    }
}

pub(super) fn reserve(bytes: &mut Vec<u8>, additional: usize) -> io::Result<()> {
    bytes.try_reserve(additional).map_err(|_| {
        io::Error::other("LEGACY_RESOURCE_LIMIT:legacy inspection source allocation failed")
    })
}
