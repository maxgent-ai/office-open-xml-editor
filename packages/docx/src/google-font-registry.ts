import {
  GOOGLE_FONT_SUBSTITUTES,
  SCRIPT_GOOGLE_FONTS,
  type FontPreloadEntry,
} from '@silurus/ooxml-core';

/** Theme-referenced typefaces commonly used by DOCX templates.
 *
 *  {@link GOOGLE_FONT_SUBSTITUTES} supplies advance-width substitutes for the
 *  base Office text faces (Calibri → Carlito, Cambria → Caladea), popular free
 *  web fonts and the Arabic Noto
 *  fallbacks — shared with pptx/xlsx. {@link SCRIPT_GOOGLE_FONTS} adds the
 *  CJK (KR/SC/TC/JP, plus HK sans) / Cyrillic / Thai / Devanagari / Hebrew
 *  Noto faces the renderer appends to the font chain. CJK fallbacks are ordered
 *  by document language. Both load only when `useGoogleFonts` is on — no binaries
 *  ship in the bundle. DOCX
 *  currently has no format-specific additions. */
export const DOCX_GOOGLE_FONTS: Record<string, FontPreloadEntry> = {
  ...GOOGLE_FONT_SUBSTITUTES,
  ...SCRIPT_GOOGLE_FONTS,
};
