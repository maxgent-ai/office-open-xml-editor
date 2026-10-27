/**
 * Font-stack primitives shared by the DOCX, PPTX and XLSX canvas font
 * builders.
 *
 * Each format keeps its own stack ORDER (which fallback classes follow the
 * authored face, and where CJK, Arabic and other script faces go): those
 * orders differ per format and are format policy. What is shared here is the
 * data every order draws from — CSS quoting, the Arabic-run test and Arabic
 * fallback faces, and the Office-face aliases derived from the vendor-neutral
 * substitution policy (`substitutions.ts`).
 */
import { quoteCssFontFamily } from '../internal/css-font-family.js';
import { ownValue } from '../internal/own-key.js';
import { OFFICE_FONT_SUBSTITUTIONS } from './substitutions.js';

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
 * Office faces whose openly licensed substitute is a DIFFERENT family (Calibri
 * → Carlito, Cambria → Caladea, the Arabic faces → Noto Arabic, …), keyed by
 * the lower-cased authored name: the vendor-neutral
 * {@link OFFICE_FONT_SUBSTITUTIONS} policy. A font stack that inserts a
 * substitute at paint time takes it from the substitutions a font source
 * actually loaded; this static alias serves only paint paths with no font
 * source session (metafile text) and script-class predicates.
 */
const OFFICE_FACE_WEB_ALIASES: Readonly<Record<string, string>> = Object.freeze(
  Object.fromEntries(OFFICE_FONT_SUBSTITUTIONS.map((entry) => [entry.for.toLowerCase(), entry.use])),
);

/** The different-family substitute for an authored Office face, if any.
 *  `family` is matched case-insensitively. */
export function officeFaceWebAlias(family: string): string | undefined {
  return ownValue(OFFICE_FACE_WEB_ALIASES, family.toLowerCase());
}

const ARABIC_SUBSTITUTE_FAMILIES = ['Noto Naskh Arabic', 'Noto Sans Arabic'];

/** Lower-cased names of the Arabic-script faces the substitution policy maps
 *  to a Noto Arabic substitute, including the Noto Arabic faces themselves. */
export const ARABIC_SUBSTITUTE_FACES: ReadonlySet<string> = new Set([
  ...OFFICE_FONT_SUBSTITUTIONS
    .filter((entry) => ARABIC_SUBSTITUTE_FAMILIES.includes(entry.use))
    .map((entry) => entry.for.toLowerCase()),
  ...ARABIC_SUBSTITUTE_FAMILIES.map((family) => family.toLowerCase()),
]);

/** The subset of {@link ARABIC_SUBSTITUTE_FACES} substituted by the
 *  Naskh-style (serif-like) Noto Naskh Arabic. */
export const NASKH_ARABIC_SUBSTITUTE_FACES: ReadonlySet<string> = new Set([
  ...OFFICE_FONT_SUBSTITUTIONS
    .filter((entry) => entry.use === 'Noto Naskh Arabic')
    .map((entry) => entry.for.toLowerCase()),
  'noto naskh arabic',
]);
