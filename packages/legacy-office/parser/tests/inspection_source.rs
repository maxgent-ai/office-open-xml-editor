//! Exercise the acquisition boundary shared by the native inspection CLIs.
#[path = "../examples/support/mod.rs"]
mod support;

#[cfg(test)]
mod tests {
    use super::support::{read_source, reserve};
    use std::io::{self, Read};

    #[test]
    fn complete_source_crosses_the_former_cli_gate_without_truncation() {
        let former_limit = 256 * 1024 * 1024;
        let source = io::repeat(0)
            .take(former_limit + 512)
            .chain(&b"complete"[..]);
        let bytes = read_source(source).unwrap();
        assert_eq!(bytes.len() as u64, former_limit + 520);
        assert_eq!(&bytes[bytes.len() - 8..], b"complete");
    }

    #[test]
    fn allocation_failure_is_classified() {
        let error = reserve(&mut vec![0], usize::MAX).unwrap_err();
        assert!(error.to_string().starts_with("LEGACY_RESOURCE_LIMIT:"));
    }

    #[test]
    fn read_failure_does_not_return_a_truncated_source() {
        struct Failed;
        impl Read for Failed {
            fn read(&mut self, _: &mut [u8]) -> io::Result<usize> {
                Err(io::Error::other("source read failed"))
            }
        }
        let error = read_source((&b"prefix"[..]).chain(Failed)).unwrap_err();
        assert_eq!(error.to_string(), "source read failed");
    }
}
