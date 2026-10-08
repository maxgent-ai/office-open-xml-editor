import { SYMBOL_MAP } from './symbol-font.js';

// GDI SYMBOL_CHARSET bytes are font codes, unlike the ordinary Unicode text
// consumed by symbol-font.ts. Adobe's Symbol Encoding to Unicode mapping:
// https://www.unicode.org/Public/MAPPINGS/VENDORS/ADOBE/symbol.txt
// Windows symbol cmap entries use the private range beginning at F020:
// https://learn.microsoft.com/en-us/typography/opentype/spec/cmap
// This bounded repertoire adds basic mathematical operators to the existing
// published Symbol markers. Unknown glyphs/fonts (including MT Extra, which
// has a different private encoding) are not guessed or decoded as ANSI.
const SYMBOL_ENCODING: Readonly<Record<number, string>> = {
  ...SYMBOL_MAP,
  0x20: ' ',
  0x2b: '+',
  0x2d: '−',
  0x3d: '=',
  0xa5: '∞',
  0xe5: '∑',
};

/** Decode a fully known font-coded string, or null for an unsupported encoding.
 *  EMF UTF-16 only remaps private F000..F0FF entries; WMF remaps its byte codes.
 *  Callers must draw decoded text with a Unicode fallback font, not Symbol. */
export function decodeMetafileSymbolText(
  text: string,
  face: string,
  privateUseOnly = false,
): string | null {
  if (face.trim().toLowerCase() !== 'symbol') return null;
  let result = '';
  for (const char of text) {
    const code = char.codePointAt(0)!;
    if (privateUseOnly && (code < 0xf000 || code > 0xf0ff)) {
      result += char;
      continue;
    }
    const target = SYMBOL_ENCODING[privateUseOnly ? code - 0xf000 : code];
    if (target === undefined) return null;
    result += target;
  }
  return result;
}
