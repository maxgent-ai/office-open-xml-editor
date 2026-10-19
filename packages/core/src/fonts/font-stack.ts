/**
 * Font-stack primitives shared by the DOCX, PPTX and XLSX canvas font
 * builders.
 *
 * Each format keeps its own stack ORDER (which fallback classes follow the
 * authored face, and where CJK, Arabic and other script faces go): those
 * orders differ per format and are format policy. What is shared here is the
 * data every order draws from — CSS quoting, the Arabic-run test and Arabic
 * fallback faces, and the Office-face web aliases derived from the single
 * Google-Fonts substitute registry.
 */
import { quoteCssFontFamily } from '../internal/css-font-family.js';
import { ownValue } from '../internal/own-key.js';
import { GOOGLE_FONT_SUBSTITUTES } from './google-fonts.js';

/** Quote each family for a CSS font-family list (no trailing comma). */
export function quoteCssFontFamilies(names: readonly string[]): string {
  return names.map(quoteCssFontFamily).join(', ');
}

/** Arabic, Arabic Supplement and Arabic Extended-A blocks. A run containing
 *  any of these takes the Arabic fallback faces. */
const ARABIC_TEXT_RE = /[؀-ۿݐ-ݿࢠ-ࣿ]/u;

/** True when `text` contains a code point that takes the Arabic fallbacks. */
export function containsArabicFallbackText(text: string): boolean {
  return ARABIC_TEXT_RE.test(text);
}

/** Arabic web fallbacks for a run's generic class: a serif face keeps the
 *  Naskh (traditional) style first; every other class uses the sans face. */
export function arabicFallbackFamilies(serif: boolean): readonly string[] {
  return serif ? ['Noto Naskh Arabic', 'Noto Sans Arabic'] : ['Noto Sans Arabic'];
}

/**
 * Office faces whose Google substitute is a DIFFERENT family (Calibri →
 * Carlito, Cambria → Caladea, the Arabic faces → Noto Arabic, …), keyed by the
 * lower-cased authored name. Entries Google serves under the requested name,
 * and the self-referencing Noto entries, are not aliases.
 */
const OFFICE_FACE_WEB_ALIASES: Readonly<Record<string, string>> = Object.freeze(
  Object.fromEntries(
    Object.entries(GOOGLE_FONT_SUBSTITUTES)
      .filter(([key, entry]) => entry.loadFamily !== undefined
        && entry.loadFamily.toLowerCase() !== key)
      .map(([key, entry]) => [key, entry.loadFamily!]),
  ),
);

/** The different-family web alias for an authored Office face, if any.
 *  `family` is matched case-insensitively. */
export function officeFaceWebAlias(family: string): string | undefined {
  return ownValue(OFFICE_FACE_WEB_ALIASES, family.toLowerCase());
}

/** Lower-cased names of the Arabic-script faces the Google registry maps to a
 *  Noto Arabic substitute, including the Noto Arabic faces themselves. */
export const ARABIC_SUBSTITUTE_FACES: ReadonlySet<string> = new Set(
  Object.entries(GOOGLE_FONT_SUBSTITUTES)
    .filter(([, entry]) => entry.loadFamily === 'Noto Naskh Arabic' || entry.loadFamily === 'Noto Sans Arabic')
    .map(([key]) => key),
);

/** The subset of {@link ARABIC_SUBSTITUTE_FACES} substituted by the
 *  Naskh-style (serif-like) Noto Naskh Arabic. */
export const NASKH_ARABIC_SUBSTITUTE_FACES: ReadonlySet<string> = new Set(
  Object.entries(GOOGLE_FONT_SUBSTITUTES)
    .filter(([, entry]) => entry.loadFamily === 'Noto Naskh Arabic')
    .map(([key]) => key),
);
