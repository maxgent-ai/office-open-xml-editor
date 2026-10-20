//! Allocation regression for OfficeArt metafile inflation: memory must follow
//! the bytes a stream produces, never its declared cbSize. A counting global
//! allocator (scoped to this test binary) records the peak live heap while
//! `inflate_exact` runs; the former `vec![0; expanded + 1]` reservation fails
//! the first assertion.
#[allow(dead_code)]
#[path = "../src/officeart/inflate.rs"]
mod inflate;

use inflate::{inflate_exact, InflateError, INFLATE_CHUNK};
use std::alloc::{GlobalAlloc, Layout, System};
use std::io::Write;
use std::sync::atomic::{AtomicUsize, Ordering};

struct Counting;

static LIVE: AtomicUsize = AtomicUsize::new(0);
static PEAK: AtomicUsize = AtomicUsize::new(0);

unsafe impl GlobalAlloc for Counting {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        let pointer = unsafe { System.alloc(layout) };
        if !pointer.is_null() {
            let live = LIVE.fetch_add(layout.size(), Ordering::SeqCst) + layout.size();
            PEAK.fetch_max(live, Ordering::SeqCst);
        }
        pointer
    }

    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        unsafe { System.dealloc(pointer, layout) };
        LIVE.fetch_sub(layout.size(), Ordering::SeqCst);
    }
}

#[global_allocator]
static ALLOCATOR: Counting = Counting;

/// MS-ODRAW caps nothing; this mirrors the reader's 32 MiB resource policy.
const CAP: usize = 32 * 1024 * 1024;

fn zlib(source: &[u8]) -> Vec<u8> {
    let mut encoder = flate2::write::ZlibEncoder::new(Vec::new(), flate2::Compression::default());
    encoder.write_all(source).unwrap();
    encoder.finish().unwrap()
}

/// Peak heap growth above the level live when `run` starts.
fn peak_growth<T>(run: impl FnOnce() -> T) -> (T, usize) {
    let base = LIVE.load(Ordering::SeqCst);
    PEAK.store(base, Ordering::SeqCst);
    let result = run();
    (result, PEAK.load(Ordering::SeqCst) - base)
}

// One test function: the counters are process-wide, so no other test may
// allocate concurrently.
#[test]
fn inflation_memory_follows_produced_bytes() {
    // A tiny stream declaring the maximum size fails without reserving it.
    let tiny = zlib(b"tiny metafile");
    let (result, growth) = peak_growth(|| inflate_exact(&tiny, CAP, tiny.len(), CAP));
    assert_eq!(result, Err(InflateError::SizeMismatch));
    assert!(
        growth < 1024 * 1024,
        "declared {CAP} bytes but peak heap grew by {growth}"
    );

    // Valid input spanning several growth steps decodes exactly, with memory
    // proportional to the output.
    let source: Vec<u8> = (0..INFLATE_CHUNK * 3 + 17)
        .map(|index| (index * 31 % 251) as u8)
        .collect();
    let stream = zlib(&source);
    let (output, growth) =
        peak_growth(|| inflate_exact(&stream, source.len(), stream.len(), CAP).unwrap());
    assert_eq!(output, source);
    assert!(growth < 3 * source.len(), "peak heap grew by {growth}");

    // Output past the declaration, or past the cap, stops one byte beyond
    // the bound and fails closed with the matching typed error.
    let bomb = zlib(&vec![0; CAP + 2 * INFLATE_CHUNK]);
    assert_eq!(
        inflate_exact(&bomb, 4096, bomb.len(), CAP),
        Err(InflateError::SizeMismatch)
    );
    let (result, growth) =
        peak_growth(|| inflate_exact(&bomb, CAP + 2 * INFLATE_CHUNK, bomb.len(), CAP));
    assert_eq!(result, Err(InflateError::OverBudget));
    assert!(growth < 3 * CAP, "peak heap grew by {growth}");
    assert_eq!(
        inflate_exact(&bomb, CAP, bomb.len(), 4096),
        Err(InflateError::OverBudget)
    );

    // Corrupt framing is distinct from size disagreements.
    let mut corrupt = stream.clone();
    corrupt[0] ^= 0xff;
    assert_eq!(
        inflate_exact(&corrupt, source.len(), corrupt.len(), CAP),
        Err(InflateError::Invalid)
    );
}
