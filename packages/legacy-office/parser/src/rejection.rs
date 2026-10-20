//! Reader admission policy, separate from document syntax and cursor misuse.
//! These internal markers are consumed only around native calls. They must not
//! be inferred from diagnostic wording in JS: limits and encryption also used
//! to share UNSUPPORTED with representational failures.

pub(crate) fn unsupported(message: impl Into<String>) -> String {
    let message = message.into();
    if message.starts_with("LEGACY_RESOURCE_LIMIT:")
        || message.starts_with("LEGACY_UNSUPPORTED_ENCRYPTION:")
        || message.starts_with("OOXML_RESOURCE_LIMIT:")
    {
        message
    } else {
        format!("UNSUPPORTED:{message}")
    }
}

pub(crate) fn resource(message: impl Into<String>) -> String {
    format!("LEGACY_RESOURCE_LIMIT:{}", message.into())
}

/// MS-DOC 2.5.2 FibBase encryption flags, MS-XLS 2.4.117 FilePass,
/// MS-PPT 2.3.2 CurrentUserAtom: these readers do not decrypt binary input.
pub(crate) fn encryption(message: impl Into<String>) -> String {
    format!("LEGACY_UNSUPPORTED_ENCRYPTION:{}", message.into())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn forwarding_preserves_resource_and_encryption_classification() {
        assert_eq!(
            unsupported("unknown structure"),
            "UNSUPPORTED:unknown structure"
        );
        let quota = resource("retained model budget");
        assert_eq!(unsupported(unsupported(quota.clone())), quota);
        let measured = "OOXML_RESOURCE_LIMIT:{\"code\":\"ooxml-resource-limit\"}";
        assert_eq!(unsupported(unsupported(measured)), measured);
        let encrypted = encryption("encrypted input");
        assert_eq!(unsupported(unsupported(encrypted.clone())), encrypted);
    }
}
