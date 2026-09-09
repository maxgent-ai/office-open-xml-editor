import {
  classifyCjkFont,
  cjkFallbackForText,
  scriptPreloadNamesForText,
  GOOGLE_FONT_SUBSTITUTES,
  SCRIPT_GOOGLE_FONTS,
  type CjkLang,
  type FontPreloadEntry,
} from '@silurus/ooxml-core';
import type { ParsedWorkbook } from './types.js';

/** Office font name → metric-compatible Google Fonts substitute for XLSX cells.
 *
 *  {@link GOOGLE_FONT_SUBSTITUTES} supplies the Office substitutes (Calibri →
 *  Carlito, Cambria → Caladea — same advance widths / vertical metrics, so
 *  text-width measurements stay close to Excel's), the popular free web fonts
 *  and the Arabic Noto fallbacks — shared with docx/pptx. {@link
 *  SCRIPT_GOOGLE_FONTS} adds the CJK / Cyrillic / Thai / Devanagari / Hebrew
 *  Noto faces (the renderer chooses the CJK Noto per cell from the cell's font
 *  name; non-CJK scripts append to the default chain). Both load only when
 *  `useGoogleFonts` is on — no binaries ship in the bundle. XLSX currently has
 *  no format-specific additions. */
export const XLSX_GOOGLE_FONTS: Record<string, FontPreloadEntry> = {
  ...GOOGLE_FONT_SUBSTITUTES,
  ...SCRIPT_GOOGLE_FONTS,
};

/** Yield every textual cell value carried by the parsed workbook: the shared
 *  string table (`text` plus rich-text `runs[].text`). This is the bulk of a
 *  workbook's painted text and is present in BOTH main and worker at parse time
 *  (sheets parse lazily, but the shared string table is workbook-level).
 *  Numbers / dates carry no script-specific glyphs, so they are irrelevant. */
function* xlsxTextRuns(wb: ParsedWorkbook | undefined): Generator<string> {
  for (const s of wb?.sharedStrings ?? []) {
    if (s.runs && s.runs.length > 0) {
      for (const r of s.runs) yield r.text;
    } else {
      yield s.text;
    }
  }
}

/**
 * The font-family names to preload for a workbook: every styled cell font, plus
 * only the script-fallback Noto faces whose script the workbook's TEXT actually
 * contains ({@link scriptPreloadNamesForText}). Office faces map to
 * metric-compatible substitutes (Calibri → Carlito, Cambria → Caladea); the
 * renderer's default chain still ends with the full Noto set, but eagerly
 * fetching the multi-MB CJK families for a workbook that has no CJK glyphs would
 * block first paint for nothing; an un-preloaded face loads lazily if it ever
 * proves needed. A workbook using only system fonts (no map entries) still
 * produces zero network requests.
 *
 * Single source of truth shared by the main-thread `_load()` and the render
 * worker. Both derive the set from the SAME parsed {@link ParsedWorkbook}, so
 * both modes preload an identical set — worker/main rendering must stay
 * pixel-equivalent.
 */
export function xlsxFontPreloadNames(wb: ParsedWorkbook | undefined, fallback?: CjkLang): Set<string> {
  const names = new Set<string>();
  let cjkLang: CjkLang | null = null;
  for (const f of wb?.styles?.fonts ?? []) {
    if (f.name) {
      names.add(f.name);
      cjkLang ??= classifyCjkFont(f.name);
    }
  }
  for (const n of scriptPreloadNamesForText(xlsxTextRuns(wb), cjkLang ?? fallback ?? null)) {
    names.add(n);
  }
  return names;
}

/** The same workbook hint is used for preloading and ambiguous cell fallback. */
export function xlsxCjkFallback(wb: ParsedWorkbook | undefined, fallback: CjkLang): CjkLang {
  for (const font of wb?.styles?.fonts ?? []) {
    const region = classifyCjkFont(font.name);
    if (region) return cjkFallbackForText(xlsxTextRuns(wb), region);
  }
  return cjkFallbackForText(xlsxTextRuns(wb), fallback);
}
