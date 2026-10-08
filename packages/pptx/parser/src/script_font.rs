//! Theme font tokens resolved per run language (issue #1627).
//!
//! A DrawingML theme token (`+mj-lt`, `+mn-ea`, `+mj-cs`, ECMA-376
//! §20.1.4.1.16-.17) names a theme font collection (major/minor), not a face.
//! PowerPoint picks the face inside that collection from the run language:
//! the collection's supplemental font for the language's script
//! (`<a:font script="Jpan" …/>`, CT_SupplementalFont) wins over the
//! collection's own latin/ea/cs face. Observed with PowerPoint's reference PDF
//! engine (#1627 controls, every face available to the engine):
//!
//! * ea: `lang` when it is an East Asian language, else `altLang`
//!   (P10, P21, A04-A06, A13); cs: `lang` when it is a complex-script
//!   language, else `altLang` (P15, P22, A07, A08, A16); latin: `lang` only
//!   (V01, V03, V04 — Vietnamese).
//! * A literal face is used as authored even when it equals a token's face
//!   (P19, P20, P24).
//! * A collection without a font for the script uses its ea/cs face
//!   (P18, I05), and an empty one leaves the slot empty: the renderer's
//!   application default then applies (N01-N05, D rows).

use std::collections::HashMap;

use crate::theme::theme_script_key;

/// Which run font slot a theme script font serves.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum FontSlot {
    Latin,
    EastAsian,
    ComplexScript,
}

/// The theme script (ECMA-376 CT_SupplementalFont `script`, ISO 15924 with
/// Office's Jpan / Hans / Hant / Hang / Uigh / Viet codes) that a BCP 47 run
/// language selects. An explicit script subtag wins; otherwise the language's
/// customary script (CLDR likely subtags) for the languages whose customary
/// script is not Latin, plus Vietnamese, whose theme entry `Viet` Office uses
/// for Latin text. Languages written in Latin or Cyrillic select no script.
pub(crate) fn language_script(lang: &str) -> Option<&'static str> {
    let mut parts = lang.split(['-', '_']).filter(|p| !p.is_empty());
    let primary = parts.next()?.to_ascii_lowercase();
    let rest: Vec<&str> = parts.collect();
    let script_subtag = rest
        .iter()
        .find(|p| p.len() == 4 && p.chars().all(|c| c.is_ascii_alphabetic()));
    if let Some(sub) = script_subtag {
        return match sub.to_ascii_lowercase().as_str() {
            "hans" => Some("Hans"),
            "hant" => Some("Hant"),
            "jpan" | "hira" | "kana" => Some("Jpan"),
            "hang" | "kore" => Some("Hang"),
            "arab" => Some(if primary == "ug" { "Uigh" } else { "Arab" }),
            "hebr" => Some("Hebr"),
            "thai" => Some("Thai"),
            "deva" => Some("Deva"),
            "beng" => Some("Beng"),
            "gujr" => Some("Gujr"),
            "guru" => Some("Guru"),
            "orya" => Some("Orya"),
            "taml" => Some("Taml"),
            "telu" => Some("Telu"),
            "knda" => Some("Knda"),
            "mlym" => Some("Mlym"),
            "sinh" => Some("Sinh"),
            "ethi" => Some("Ethi"),
            "tibt" => Some("Tibt"),
            "mong" => Some("Mong"),
            "yiii" => Some("Yiii"),
            "mymr" => Some("Mymr"),
            "khmr" => Some("Khmr"),
            "laoo" => Some("Laoo"),
            "armn" => Some("Armn"),
            "geor" => Some("Geor"),
            "cher" => Some("Cher"),
            "syrc" => Some("Syrc"),
            "thaa" => Some("Thaa"),
            "nkoo" => Some("Nkoo"),
            _ => None,
        };
    }
    let region = rest
        .iter()
        .find(|p| p.len() == 2 && p.chars().all(|c| c.is_ascii_alphabetic()))
        .map(|r| r.to_ascii_uppercase());
    Some(match primary.as_str() {
        "ja" => "Jpan",
        "ko" => "Hang",
        "zh" => match region.as_deref() {
            Some("TW" | "HK" | "MO") => "Hant",
            _ => "Hans",
        },
        "he" | "yi" => "Hebr",
        "ar" | "fa" | "ur" | "ps" | "sd" => "Arab",
        "ug" => "Uigh",
        "th" => "Thai",
        "hi" | "mr" | "ne" | "sa" | "kok" => "Deva",
        "bn" | "as" => "Beng",
        "gu" => "Gujr",
        "pa" => "Guru",
        "or" => "Orya",
        "ta" => "Taml",
        "te" => "Telu",
        "kn" => "Knda",
        "ml" => "Mlym",
        "si" => "Sinh",
        "am" | "ti" => "Ethi",
        "bo" | "dz" => "Tibt",
        "ii" => "Yiii",
        "my" => "Mymr",
        "km" => "Khmr",
        "lo" => "Laoo",
        "hy" => "Armn",
        "ka" => "Geor",
        "chr" => "Cher",
        "syr" => "Syrc",
        "dv" => "Thaa",
        "vi" => "Viet",
        _ => return None,
    })
}

/// The run slot a theme script font serves. Jpan/Hans/Hant/Hang serve the East
/// Asian slot. Viet and Cher serve the Latin slot: under vi-VN / chr-Cher-US
/// the latin token took the Viet / Cher font (#1627 L37, L40, V01, V04), and
/// Cherokee characters are drawn in the latin face (S rows). Every other
/// script serves the complex-script slot.
pub(crate) fn script_slot(script: &str) -> FontSlot {
    match script {
        "Jpan" | "Hans" | "Hant" | "Hang" => FontSlot::EastAsian,
        "Viet" | "Cher" => FontSlot::Latin,
        _ => FontSlot::ComplexScript,
    }
}

/// The theme script that selects a `slot` face for a run: the language's
/// script when it serves that slot, else (ea and cs only) the alternate
/// language's.
pub(crate) fn slot_script(
    slot: FontSlot,
    lang: Option<&str>,
    alt_lang: Option<&str>,
) -> Option<&'static str> {
    let serves = |l: Option<&str>| {
        l.and_then(language_script)
            .filter(|script| script_slot(script) == slot)
    };
    match slot {
        FontSlot::Latin => serves(lang),
        FontSlot::EastAsian | FontSlot::ComplexScript => serves(lang).or_else(|| serves(alt_lang)),
    }
}

/// Resolve an authored typeface for `slot`. A theme token (`+mj-*` / `+mn-*`)
/// picks its collection; inside it the language's script font, else the face
/// the token names. `None` = the slot has no face (the renderer
/// applies Office's application default). A literal face is returned as is;
/// an empty literal is `None` for ea/cs.
pub(crate) fn resolve_slot_face(
    authored: &str,
    slot: FontSlot,
    theme: &HashMap<String, String>,
    lang: Option<&str>,
    alt_lang: Option<&str>,
) -> Option<String> {
    if authored.is_empty() {
        return None;
    }
    let Some(set) = theme_token_set(authored) else {
        return Some(authored.to_owned());
    };
    if let Some(face) = slot_script(slot, lang, alt_lang)
        .and_then(|script| theme.get(&theme_script_key(set, script)))
        .filter(|face| !face.is_empty())
    {
        return Some(face.clone());
    }
    theme.get(authored).filter(|face| !face.is_empty()).cloned()
}

/// `+mj` / `+mn` for a theme font token (`+mj-lt`, `+mn-ea`, …).
pub(crate) fn theme_token_set(typeface: &str) -> Option<&'static str> {
    if typeface.starts_with("+mj-") {
        Some("+mj")
    } else if typeface.starts_with("+mn-") {
        Some("+mn")
    } else {
        None
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn theme() -> HashMap<String, String> {
        [
            ("+mn-lt", "Corbel"),
            ("+mn-ea", "Meiryo"),
            ("+mn-cs", "Microsoft Sans Serif"),
            ("+mn-script-Jpan", "Yu Mincho"),
            ("+mn-script-Hebr", "David"),
            ("+mn-script-Viet", "Palatino Linotype"),
            ("+mj-lt", "Garamond"),
            ("+mj-ea", "HGMinchoE"),
            ("+mj-script-Jpan", "HGSoeiKakugothicUB"),
        ]
        .into_iter()
        .map(|(k, v)| (k.to_owned(), v.to_owned()))
        .collect()
    }

    #[test]
    fn language_scripts_follow_region_and_script_subtags() {
        assert_eq!(language_script("ja-JP"), Some("Jpan"));
        assert_eq!(language_script("ja"), Some("Jpan"));
        assert_eq!(language_script("zh-CN"), Some("Hans"));
        assert_eq!(language_script("zh-SG"), Some("Hans"));
        assert_eq!(language_script("zh-TW"), Some("Hant"));
        assert_eq!(language_script("zh-HK"), Some("Hant"));
        assert_eq!(language_script("zh-MO"), Some("Hant"));
        assert_eq!(language_script("mn-Mong-CN"), Some("Mong"));
        assert_eq!(language_script("chr-Cher-US"), Some("Cher"));
        assert_eq!(language_script("yi-001"), Some("Hebr"));
        assert_eq!(language_script("ug-CN"), Some("Uigh"));
        assert_eq!(language_script("en-US"), None);
        assert_eq!(language_script("fr-FR"), None);
        assert_eq!(language_script("mn-MN"), None);
    }

    #[test]
    fn east_asian_and_complex_script_use_lang_then_alt_lang_latin_uses_lang_only() {
        use FontSlot::*;
        assert_eq!(
            slot_script(EastAsian, Some("ko-KR"), Some("ja-JP")),
            Some("Hang")
        );
        assert_eq!(
            slot_script(EastAsian, Some("he-IL"), Some("ja-JP")),
            Some("Jpan")
        );
        assert_eq!(
            slot_script(ComplexScript, Some("ja-JP"), Some("he-IL")),
            Some("Hebr")
        );
        assert_eq!(
            slot_script(ComplexScript, Some("ar-SA"), Some("he-IL")),
            Some("Arab")
        );
        assert_eq!(slot_script(Latin, Some("vi-VN"), None), Some("Viet"));
        assert_eq!(slot_script(Latin, Some("en-US"), Some("vi-VN")), None);
    }

    #[test]
    fn tokens_pick_the_script_font_of_their_collection_literals_stay() {
        let t = theme();
        use FontSlot::*;
        let r = |a: &str, s, l: Option<&str>, al: Option<&str>| resolve_slot_face(a, s, &t, l, al);
        assert_eq!(
            r("+mn-ea", EastAsian, Some("ja-JP"), None).as_deref(),
            Some("Yu Mincho")
        );
        assert_eq!(
            r("+mn-ea", EastAsian, Some("en-US"), None).as_deref(),
            Some("Meiryo")
        );
        assert_eq!(
            r("+mj-ea", EastAsian, Some("ja-JP"), None).as_deref(),
            Some("HGSoeiKakugothicUB")
        );
        assert_eq!(
            r("Meiryo", EastAsian, Some("ja-JP"), None).as_deref(),
            Some("Meiryo")
        );
        assert_eq!(
            r("+mn-cs", ComplexScript, Some("he-IL"), None).as_deref(),
            Some("David")
        );
        assert_eq!(r("+mj-cs", ComplexScript, Some("he-IL"), None), None);
        assert_eq!(
            r("+mn-lt", Latin, Some("vi-VN"), None).as_deref(),
            Some("Palatino Linotype")
        );
        assert_eq!(
            r("+mj-lt", Latin, Some("vi-VN"), None).as_deref(),
            Some("Garamond")
        );
        assert_eq!(r("", EastAsian, Some("ja-JP"), None), None);
    }
}
