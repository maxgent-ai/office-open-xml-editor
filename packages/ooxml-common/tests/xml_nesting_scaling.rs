use ooxml_common::depth::xml_nesting_exceeds;
use std::time::{Duration, Instant};

fn elapsed(xml: &[u8]) -> Duration {
    let start = Instant::now();
    // Force the complete scan rather than timing the early rejection prefix.
    assert!(!xml_nesting_exceeds(std::hint::black_box(xml), u32::MAX));
    start.elapsed()
}

#[test]
fn lexical_scanning_scales_linearly_for_twenty_and_eighty_thousand_units() {
    for kind in ["nesting", "attributes", "declarations-and-siblings"] {
        let input = |size| match kind {
            "nesting" => "<a xmlns:p='u'>".repeat(size) + &"</a>".repeat(size),
            "attributes" => format!("<a{}></a>", " a='/> < >' b=\"< > /\"".repeat(size)),
            _ => format!(
                "<!DOCTYPE r [{}]><r>{}</r>",
                "<!ENTITY x '<a> ] >'><!-- ] > --><?pi ] > ?>".repeat(size),
                "<a x='>' /><b><!-- <x> --><![CDATA[ <x> ]]><?pi <x> ?></b>".repeat(size)
            ),
        };
        let small = input(20_000);
        let large = input(80_000);
        elapsed(small.as_bytes());
        elapsed(large.as_bytes());
        let mut small_times = Vec::new();
        let mut large_times = Vec::new();
        for _ in 0..5 {
            small_times.push(elapsed(small.as_bytes()));
            large_times.push(elapsed(large.as_bytes()));
        }
        small_times.sort();
        large_times.sort();
        let ratio = large_times[2].as_secs_f64() / small_times[2].as_secs_f64();
        eprintln!(
            "{kind}: {:?} -> {:?}, ratio {ratio:.2}",
            small_times[2], large_times[2]
        );
        // Four times the input should take roughly four times the work. This
        // leaves scheduler headroom while rejecting quadratic (16x) rescanning.
        assert!(ratio < 10.0, "{kind}: 4x input took {ratio:.2}x time");
    }
}
