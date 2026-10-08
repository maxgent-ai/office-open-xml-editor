import { SYMBOL_MAP } from './symbol-font.js';

// GDI SYMBOL_CHARSET bytes are font codes, unlike the ordinary Unicode text
// consumed by symbol-font.ts. Adobe's Symbol Encoding to Unicode mapping:
// https://www.unicode.org/Public/MAPPINGS/VENDORS/ADOBE/symbol.txt
// Windows symbol cmap entries use the private range beginning at F020:
// https://learn.microsoft.com/en-us/typography/opentype/spec/cmap
// This bounded repertoire adds basic mathematical operators to the existing
// published Symbol markers. Unknown glyphs/fonts are not guessed or decoded
// as ANSI; each exact selected family has its own published repertoire.
const SYMBOL_ENCODING: Readonly<Record<number, string>> = {
  ...SYMBOL_MAP,
  0x20: ' ',
  0x2b: '+',
  0x2d: '−',
  0x3d: '=',
  0xa5: '∞',
  0xe5: '∑',
};

// Mozilla's published MT Extra encoding (the reduced Design Science face):
// https://www-archive.mozilla.org/projects/mathml/fonts/encoding/mtextra
// The five ellipsis positions also agree with Design Science 4.31's enhanced
// font cmap and outlines. This is a bounded font-code conversion, not a claim
// that every same-named version shares the full repertoire: older/enhanced
// versions differ elsewhere. No fonts or font-specific glyph indices are
// embedded. Unknown entries still refuse the entire string atomically.
const MT_EXTRA_ENCODING: Readonly<Record<number, string>> = {
  0x20: ' ',
  0x4b: '…', // U+2026 horizontal (baseline) ellipsis
  0x4c: '⋯', // U+22EF math-axis ellipsis, distinct from U+2026
  0x4d: '⋮', // U+22EE vertical ellipsis
  0x4e: '⋰', // U+22F0 up-right diagonal ellipsis
  0x4f: '⋱', // U+22F1 down-right diagonal ellipsis
};

/** Decode a fully known font-coded string, or null for an unsupported encoding.
 *  EMF UTF-16 only remaps private F000..F0FF entries; WMF remaps its byte codes.
 *  Callers must draw decoded text with a Unicode fallback font, not the source
 *  symbol font (which could reinterpret the new Unicode points). */
export function decodeMetafileSymbolText(
  text: string,
  face: string,
  privateUseOnly = false,
): string | null {
  const family = face.trim().toLowerCase();
  const table = family === 'symbol' ? SYMBOL_ENCODING
    : family === 'mt extra' ? MT_EXTRA_ENCODING : null;
  if (!table) return null;
  let result = '';
  for (const char of text) {
    const code = char.codePointAt(0)!;
    if (privateUseOnly && (code < 0xf000 || code > 0xf0ff)) {
      result += char;
      continue;
    }
    const target = table[privateUseOnly ? code - 0xf000 : code];
    if (target === undefined) return null;
    result += target;
  }
  return result;
}
