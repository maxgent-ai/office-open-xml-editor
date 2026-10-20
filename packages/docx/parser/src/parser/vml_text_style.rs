use super::parse_vml_length_pt;

#[derive(Default)]
pub(super) struct Font<'a> {
    pub family: Option<&'a str>,
    pub size_pt: Option<f64>,
    pub bold: Option<bool>,
    pub italic: Option<bool>,
}

/// Part 4 §19.1.2.23 uses CSS2 `font`: expand shorthand resets within each
/// declaration layer, in source order, before applying shape over shapetype.
/// An omitted shorthand weight/style means normal, not the template's value.
/// Keep the family suffix intact; whitespace in a family is not a font boundary.
pub(super) fn resolve<'a>(instance: &'a str, template: &'a str) -> Font<'a> {
    let base = layer(template);
    let top = layer(instance);
    Font {
        family: top.family.or(base.family),
        size_pt: top.size_pt.or(base.size_pt),
        bold: top.bold.or(base.bold),
        italic: top.italic.or(base.italic),
    }
}

fn family(value: &str) -> Option<&str> {
    let value = value.trim();
    // Strip one enclosing quote pair only for a single quoted family. A CSS
    // family list must retain its delimiters and quotes for the font authority.
    let value = if let Some(quote @ ('\'' | '"')) = value.chars().next() {
        let tail = &value[1..];
        if tail.find(quote) == Some(tail.len().saturating_sub(1)) {
            &tail[..tail.len() - 1]
        } else {
            value
        }
    } else {
        value
    };
    (!value.is_empty()).then_some(value)
}

fn weight(value: &str) -> bool {
    value.eq_ignore_ascii_case("bold")
        || value.eq_ignore_ascii_case("bolder")
        || value.parse::<u32>().is_ok_and(|n| n >= 500)
}

fn slant(value: &str) -> bool {
    value.eq_ignore_ascii_case("italic") || value.eq_ignore_ascii_case("oblique")
}

fn token(raw: &str) -> (&str, &str) {
    let raw = raw.trim_start();
    let end = raw
        .find(|c: char| c.is_whitespace() || c == '/')
        .unwrap_or(raw.len());
    (&raw[..end], raw[end..].trim_start())
}

fn shorthand(mut raw: &str) -> Option<Font<'_>> {
    let mut font = Font {
        bold: Some(false),
        italic: Some(false),
        ..Default::default()
    };
    loop {
        let (part, rest) = token(raw);
        let lower = part.to_ascii_lowercase();
        let absolute = ["pt", "px", "in", "cm", "mm", "pc"]
            .iter()
            .any(|unit| lower.ends_with(unit));
        let relative = ["em", "ex", "%"].iter().any(|unit| lower.ends_with(unit))
            || matches!(
                lower.as_str(),
                "xx-small"
                    | "x-small"
                    | "small"
                    | "medium"
                    | "large"
                    | "x-large"
                    | "xx-large"
                    | "smaller"
                    | "larger"
            );
        if absolute || relative || lower == "0" {
            if absolute || lower == "0" {
                font.size_pt = parse_vml_length_pt(part).filter(|n| n.is_finite() && *n >= 0.0);
                font.size_pt?;
            }
            // Relative/keyword sizes have no acquired point metric yet. Keep
            // their family/style, which existing fitted text already consumes;
            // do not discard those independent facts or invent a point size.
            // Line height belongs to the compound grammar, not the family.
            // Fitted VML glyph placement uses the retained path/shape extent.
            let rest = if let Some(height) = rest.strip_prefix('/') {
                let (height, rest) = token(height);
                if height.is_empty() {
                    return None;
                }
                rest
            } else {
                rest
            };
            font.family = family(rest);
            font.family?;
            return Some(font);
        }
        if part.is_empty() {
            return None;
        }
        match lower.as_str() {
            "italic" | "oblique" => font.italic = Some(true),
            "bold" | "bolder" | "lighter" => font.bold = Some(weight(&lower)),
            "normal" | "small-caps" => {}
            _ if lower.parse::<u32>().is_ok() => font.bold = Some(weight(&lower)),
            _ => return None,
        }
        raw = rest;
    }
}

fn layer(style: &str) -> Font<'_> {
    let mut font = Font::default();
    let mut quote = None;
    let mut escaped = false;
    // CSS strings can contain semicolons; they do not end a declaration.
    for declaration in style.split(|c| {
        if escaped {
            escaped = false;
            return false;
        }
        if c == '\\' {
            escaped = true;
            return false;
        }
        if Some(c) == quote {
            quote = None;
        } else if quote.is_none() && matches!(c, '\'' | '"') {
            quote = Some(c);
        }
        c == ';' && quote.is_none()
    }) {
        let Some((key, value)) = declaration.split_once(':') else {
            continue;
        };
        let value = value.trim();
        match key.trim().to_ascii_lowercase().as_str() {
            "font" => {
                if let Some(expanded) = shorthand(value) {
                    font = expanded;
                }
            }
            "font-family" => font.family = family(value).or(font.family),
            "font-size" => {
                font.size_pt = parse_vml_length_pt(value)
                    .filter(|n| n.is_finite() && *n >= 0.0)
                    .or(font.size_pt)
            }
            "font-weight" => font.bold = Some(weight(value)),
            "font-style" => font.italic = Some(slant(value)),
            _ => {}
        }
    }
    font
}
