// Shared package inputs exercise each format's own optional-part error path.
pub fn lexical_depth_cases() -> [(&'static str, String, bool); 3] {
    let deep = "<a>".repeat(257) + &"</a>".repeat(257);
    [
        ("well-formed deep part", deep.clone(), true),
        (
            "malformed deep part",
            format!("<r a='1' a='2'>{deep}</r>"),
            true,
        ),
        (
            "shallow malformed siblings",
            format!(
                "<styles>{}</styles>",
                "<style a='> /' b=\"/>\" a='2' />".repeat(257)
            ),
            false,
        ),
    ]
}
