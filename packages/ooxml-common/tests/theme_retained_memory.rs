use ooxml_common::theme::ThemeFormatScheme;
use std::alloc::{GlobalAlloc, Layout, System};
use std::sync::atomic::{AtomicUsize, Ordering};

struct CountingAllocator;
static LIVE_BYTES: AtomicUsize = AtomicUsize::new(0);

unsafe impl GlobalAlloc for CountingAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        let ptr = unsafe { System.alloc(layout) };
        if !ptr.is_null() {
            LIVE_BYTES.fetch_add(layout.size(), Ordering::Relaxed);
        }
        ptr
    }

    unsafe fn dealloc(&self, ptr: *mut u8, layout: Layout) {
        LIVE_BYTES.fetch_sub(layout.size(), Ordering::Relaxed);
        unsafe { System.dealloc(ptr, layout) };
    }

    unsafe fn realloc(&self, ptr: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        let replacement = unsafe { System.realloc(ptr, layout, new_size) };
        if !replacement.is_null() {
            LIVE_BYTES.fetch_add(new_size, Ordering::Relaxed);
            LIVE_BYTES.fetch_sub(layout.size(), Ordering::Relaxed);
        }
        replacement
    }
}

#[global_allocator]
static ALLOCATOR: CountingAllocator = CountingAllocator;

fn theme(ancestor_namespaces: usize, line_styles: usize) -> String {
    const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
    let mut xml = format!("<a:theme xmlns:a=\"{A}\" ");
    for i in 0..ancestor_namespaces {
        xml.push_str(&format!("xmlns:u{i}=\"urn:unused:{i}\" "));
    }
    xml.push_str("><a:themeElements><a:fmtScheme name=\"memory\"><a:lnStyleLst>");
    for i in 0..line_styles {
        xml.push_str(&format!("<a:ln xmlns:local=\"urn:entry:{i}\" w=\"25400\"><a:solidFill><a:srgbClr val=\"123456\"/></a:solidFill></a:ln>"));
    }
    xml.push_str("</a:lnStyleLst></a:fmtScheme></a:themeElements></a:theme>");
    xml
}

fn retained_bytes(xml: &str) -> usize {
    let before = LIVE_BYTES.load(Ordering::Relaxed);
    let scheme = ThemeFormatScheme::parse(xml);
    std::hint::black_box(&scheme);
    let retained = LIVE_BYTES.load(Ordering::Relaxed) - before;
    drop(scheme);
    retained
}

#[test]
fn local_style_declarations_do_not_multiply_retained_ancestor_scopes() {
    // One integration test owns the counting allocator, avoiding test-thread noise.
    // These are the review reproducer's shapes: every style has a distinct local
    // declaration while a large, unused scope is inherited from the theme.
    let small = theme(512, 2048);
    let large = theme(1024, 4096);
    let small_retained = retained_bytes(&small);
    let large_retained = retained_bytes(&large);
    assert!(
        small_retained <= small.len() * 8,
        "small retained {small_retained} bytes for {} input bytes",
        small.len()
    );
    assert!(
        large_retained <= large.len() * 8,
        "large retained {large_retained} bytes for {} input bytes",
        large.len()
    );
    assert!(
        large_retained <= small_retained * 3,
        "retained bytes grew from {small_retained} to {large_retained}"
    );
}
