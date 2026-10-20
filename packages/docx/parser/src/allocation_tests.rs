//! Thread-local live-allocation peaks for synchronous parser regression tests.
use std::alloc::{GlobalAlloc, Layout, System};
use std::cell::Cell;

struct CountingAllocator;
thread_local! {
    static COUNTS: Cell<Option<(isize, usize)>> = const { Cell::new(None) };
}

fn account(delta: isize) {
    let _ = COUNTS.try_with(|counts| {
        if let Some((live, peak)) = counts.get() {
            let live = live + delta;
            counts.set(Some((live, peak.max(live.max(0) as usize))));
        }
    });
}

unsafe impl GlobalAlloc for CountingAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        let ptr = unsafe { System.alloc(layout) };
        if !ptr.is_null() {
            account(layout.size() as isize);
        }
        ptr
    }

    unsafe fn dealloc(&self, ptr: *mut u8, layout: Layout) {
        account(-(layout.size() as isize));
        unsafe { System.dealloc(ptr, layout) };
    }

    unsafe fn realloc(&self, ptr: *mut u8, layout: Layout, size: usize) -> *mut u8 {
        let ptr = unsafe { System.realloc(ptr, layout, size) };
        if !ptr.is_null() {
            account(size as isize - layout.size() as isize);
        }
        ptr
    }
}

#[global_allocator]
static ALLOCATOR: CountingAllocator = CountingAllocator;

pub(crate) fn peak(f: impl FnOnce()) -> usize {
    struct Reset;
    impl Drop for Reset {
        fn drop(&mut self) {
            COUNTS.with(|counts| counts.set(None));
        }
    }
    COUNTS.with(|counts| {
        assert!(counts.get().is_none());
        counts.set(Some((0, 0)));
    });
    let _reset = Reset;
    f();
    COUNTS.with(|counts| counts.get().unwrap().1)
}
