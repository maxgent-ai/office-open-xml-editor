import { classifyCjkFont, findReferenceFontMetrics } from '@silurus/ooxml-core';
import {
  EAST_ASIAN_SLOT_RANGES,
  MICROSOFT_JHENGHEI_RANGES,
  MS_MINCHO_GOTHIC_BITMAP_BASE64,
  PMINGLIU_RANGES,
} from './east-asian-default-coverage.js';

// PowerPoint's application defaults for a run whose East Asian or complex-
// script slot has no face (issue #1627). The parser resolves every authored
// ea/cs face, including theme tokens and the language's theme script font;
// these apply only when that leaves the slot empty: no ea/cs anywhere in the
// chain, an explicit typeface="", or a token naming an empty theme slot.
//
// Observed with PowerPoint's reference PDF engine (#1627 controls D, N, X):
// * A Latin face that is itself an East Asian face (OS/2 Far-East code page:
//   Yu Gothic, SimSun, Batang, PMingLiU, MS Mincho) draws the East Asian text.
//   Characters it lacks fell back to Microsoft JhengHei / Malgun Gothic after
//   a Japanese face (Yu Gothic, MS Mincho) and to PMingLiU / Batang after a
//   Chinese or Korean face (SimSun, Batang, PMingLiU).
// * Otherwise the Latin face's installed PANOSE serif style picks the tier:
//   serif (2-10: Perpetua, Garamond, Georgia) MS Mincho > PMingLiU > Batang,
//   sans (Corbel, Arial, Tahoma, Gill Sans MT, and any face not installed)
//   MS Gothic > Microsoft JhengHei > Malgun Gothic. Authored panose /
//   pitchFamily on <a:latin> has no effect.
// * The first tier face whose repertoire covers the whole East Asian text
//   draws it: Japanese and Traditional Chinese in MS Mincho / MS Gothic,
//   "简体中文" entirely in PMingLiU / Microsoft JhengHei, Hangul in Batang /
//   Malgun Gothic.
// * Complex scripts: Hebrew and Arabic in Arial, Thai in Angsana New,
//   Devanagari in Mangal, Tamil in Latha, whatever the Latin face. The same
//   faces follow an authored cs face that lacks a character.

const SERIF_TIERS = ['MS Mincho', 'PMingLiU', 'Batang'] as const;
const SANS_TIERS = ['MS Gothic', 'Microsoft JhengHei', 'Malgun Gothic'] as const;

/** Complex-script application defaults, in per-glyph fallback order. */
export const COMPLEX_SCRIPT_DEFAULT_FACES: readonly string[] = Object.freeze([
  'Arial', 'Angsana New', 'Mangal', 'Latha',
]);

const THAI_RE = /\p{Script=Thai}/u;
const DEVANAGARI_RE = /\p{Script=Devanagari}/u;
const TAMIL_RE = /\p{Script=Tamil}/u;

/** The complex-script application default for one character. */
export function complexScriptDefaultFace(ch: string): string {
  if (THAI_RE.test(ch)) return 'Angsana New';
  if (DEVANAGARI_RE.test(ch)) return 'Mangal';
  if (TAMIL_RE.test(ch)) return 'Latha';
  return 'Arial';
}

let minchoBits: Uint8Array | undefined;

function decodeBitmap(): Uint8Array {
  if (minchoBits) return minchoBits;
  const binary = atob(MS_MINCHO_GOTHIC_BITMAP_BASE64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  minchoBits = bytes;
  return bytes;
}

function slotIndex(cp: number): number {
  let offset = 0;
  for (const [lo, hi] of EAST_ASIAN_SLOT_RANGES) {
    if (cp >= lo && cp <= hi) return offset + cp - lo;
    offset += hi - lo + 1;
  }
  return -1;
}

function inRanges(cp: number, ranges: readonly (readonly [number, number])[]): boolean {
  let lo = 0;
  let hi = ranges.length - 1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    const [a, b] = ranges[mid];
    if (cp < a) hi = mid - 1;
    else if (cp > b) lo = mid + 1;
    else return true;
  }
  return false;
}

function coveredByMincho(cp: number): boolean {
  const i = slotIndex(cp);
  if (i < 0) return false;
  return (decodeBitmap()[i >> 3] & (1 << (i & 7))) !== 0;
}

function covers(face: string, text: string): boolean {
  for (const ch of text) {
    const cp = ch.codePointAt(0) ?? 0;
    if (/\s/u.test(ch)) continue;
    const ok = face === 'MS Mincho' || face === 'MS Gothic' ? coveredByMincho(cp)
      : face === 'PMingLiU' ? inRanges(cp, PMINGLIU_RANGES)
      : face === 'Microsoft JhengHei' ? inRanges(cp, MICROSOFT_JHENGHEI_RANGES)
      : false;
    if (!ok) return false;
  }
  return true;
}

function profilesOf(face: string) {
  const office = findReferenceFontMetrics(face, { source: 'office-mac' });
  return office.length > 0 ? office : findReferenceFontMetrics(face);
}

/** True when the installed face declares an OS/2 Far-East code page. */
export function isEastAsianFace(face: string): boolean {
  return profilesOf(face).some((p) => p.farEastCodePage === true);
}

/** PANOSE serif class of the installed face; unknown faces are sans. */
export function isSerifLatinFace(face: string): boolean {
  return profilesOf(face).some((p) => {
    const panose = p.panose;
    return !!panose && panose[0] === 2 && panose[1] >= 2 && panose[1] <= 10;
  });
}

/**
 * Faces, in fallback order, for East Asian text drawn with an empty East Asian
 * slot. `latinFace` is the run's resolved Latin face; `text` the run's East
 * Asian-slot characters.
 */
export function eastAsianDefaultFaces(latinFace: string | null, text: string): readonly string[] {
  if (latinFace && isEastAsianFace(latinFace)) {
    return classifyCjkFont(latinFace) === 'jp'
      ? [latinFace, 'Microsoft JhengHei', 'Malgun Gothic']
      : [latinFace, 'PMingLiU', 'Batang'];
  }
  const tiers = latinFace && isSerifLatinFace(latinFace) ? SERIF_TIERS : SANS_TIERS;
  const first = tiers.slice(0, 2).find((face) => covers(face, text));
  return first ? [first, ...tiers.filter((face) => face !== first)] : [tiers[2], tiers[0], tiers[1]];
}
