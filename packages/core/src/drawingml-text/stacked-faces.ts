import { findReferenceFontMetrics } from '../fonts/reference-font-metrics.js';

/**
 * Face and character data for stacked vertical text, `a:bodyPr@vert`
 * `wordArtVert` / `wordArtVertRtl` (ECMA-376 §20.1.10.83 ST_TextVerticalType:
 * "one letter on top of another"). The standard defines no metrics; every
 * value below is observed PowerPoint behaviour (issue #1626 control decks,
 * PowerPoint for Mac 16.x, reference "electronic distribution" PDF export).
 *
 * Cell. Each upright glyph takes a cell of 7/6 × the face box × size along
 * the column, where the face box is usWinAscent + usWinDescent, or
 * sTypoAscender − sTypoDescender + sTypoLineGap for a face that sets
 * USE_TYPO_METRICS (the #1610 glyph box). The rule fits 58 of 63 measured
 * faces (12 to 48 pt, regular and bold) to 0.03 %. Five face families do not follow any
 * font-table quantity (usWin, hhea with or without lineGap, typo, head bbox,
 * OS/2 version); their measured boxes are in STACKED_CELL_BOX_OVERRIDES.
 */

/** 7/6: the ratio of a stacked cell to the face box. */
export const STACKED_CELL_FACTOR = 7 / 6;

function normalizeFamily(family: string): string {
  return family.normalize('NFKC').trim().replace(/\s+/gu, ' ').toLocaleLowerCase('en-US');
}

/**
 * Canonical (reference-catalog) family for `family`, lower-cased, so a
 * localized or PostScript alias (游ゴシック, MS-Gothic, 宋体) finds the same
 * table entry as the English name.
 */
function canonicalFamily(family: string): string {
  const profiles = findReferenceFontMetrics(family);
  return normalizeFamily(profiles.length > 0 ? profiles[0].family : family);
}

/**
 * Measured face boxes (em, before the 7/6 cell factor) for the families whose
 * stacked cell is not their font-table box. Controls: wordartvert-supp
 * sandwich columns (20 pt) and wordartvert-supp2 cell columns (12, 18, 24, 32,
 * 48 pt, regular and bold), glyph pitch corrected for the export's 1.0015
 * vertical text scale.
 * - Yu Gothic Regular / Light / Medium / Bold: 1.5084 (usWin box 1.2871).
 *   Yu Mincho, with the same usWin/hhea/typo values, follows its usWin box.
 * - SimSun, NSimSun: 1.0380 (usWin 1.0). SimHei, KaiTi and FangSong carry the
 *   same tables and follow usWin.
 * - MingLiU, PMingLiU, MingLiU_HKSCS: 1.1132; MingLiU-ExtB and
 *   PMingLiU-ExtB: 1.1137 (usWin 1.0).
 * - Gulim, Dotum: 1.0481 (usWin 1.0); Batang, BatangChe and Gungsuh follow usWin.
 * - STZhongsong: 1.2165 (usWin 1.137).
 */
const STACKED_CELL_BOX_OVERRIDES: ReadonlyMap<string, number> = new Map([
  ['yu gothic', 1.5084],
  ['yu gothic medium', 1.5084],
  ['simsun', 1.038],
  ['nsimsun', 1.038],
  ['mingliu', 1.1132],
  ['pmingliu', 1.1132],
  ['mingliu_hkscs', 1.1132],
  ['mingliu-extb', 1.1137],
  ['pmingliu-extb', 1.1137],
  ['gulim', 1.0481],
  ['dotum', 1.0481],
  ['stzhongsong', 1.2165],
]);

/**
 * The font-table face box (em) of `family`: usWin, or typo under
 * USE_TYPO_METRICS. A named weight such as "Yu Mincho Light" is catalogued
 * under its own weight, so a lookup at the run's weight that finds nothing
 * falls back to any weight of the family. Undefined for an unknown face.
 */
export function stackedFaceBox(
  family: string,
  bold: boolean,
  italic: boolean,
): { ascent: number; descent: number } | undefined {
  let profiles = findReferenceFontMetrics(family, { weight: bold ? 700 : 400, style: italic ? 'italic' : 'normal' });
  if (profiles.length === 0) profiles = findReferenceFontMetrics(family);
  const profile = profiles.find((p) => p.source === 'macos-supplemental')
    ?? profiles.find((p) => p.source === 'office-mac') ?? profiles[0];
  if (!profile || !(profile.unitsPerEm > 0)) return undefined;
  if (profile.typoMetrics) {
    const [ascender, descender, lineGap] = profile.typoMetrics;
    return { ascent: (ascender + Math.max(0, lineGap)) / profile.unitsPerEm, descent: -descender / profile.unitsPerEm };
  }
  if (!profile.win) return undefined;
  return { ascent: profile.win[0] / profile.unitsPerEm, descent: profile.win[1] / profile.unitsPerEm };
}

/** The measured face box of `family` (em), or undefined to use its font tables. */
export function stackedCellBoxOverride(family: string): number | undefined {
  return STACKED_CELL_BOX_OVERRIDES.get(canonicalFamily(family))
    ?? STACKED_CELL_BOX_OVERRIDES.get(normalizeFamily(family));
}

/**
 * Characters PowerPoint turns sideways in stacked text. It is not UAX #50: of
 * about 150 characters measured (ASCII and Latin-1 punctuation, general
 * punctuation, arrows, CJK punctuation, brackets, fullwidth forms, small kana),
 * exactly these are handled differently; every other character stands upright
 * in its cell, including ，．：；？！, small kana, 〜, – — and « ».
 * Characters outside the measured set are treated as upright.
 */
const STACKED_SIDEWAYS = new Set(
  [...'()<>[]{}…‥‐←→↑↓―∥、。ー～＝＿￣｜（）〔〕［］｛｝〈〉《》「」『』【】〖〗'].map((c) => c.codePointAt(0)!),
);

export function stackedSidewaysCharacter(cp: number): boolean {
  return STACKED_SIDEWAYS.has(cp);
}

/**
 * Which sideways characters a face draws with its own vertical glyph. Such a
 * glyph stands upright in a 1 em cell; a face without one draws the character
 * rotated 90° clockwise. Controls: Yu Gothic, MS Gothic, Meiryo and SimSun
 * draw their vertical glyphs; Arial and Cambria Math rotate. The sets are the
 * members of each Office-bundled face's GSUB `vert` coverage. Yu Gothic also
 * draws ∥ with a vertical glyph (wordartvert-supp slide 7), which its `vert`
 * lookup does not list; the other faces follow their `vert` coverage only.
 * A face not in this table (every Latin face) has no vertical glyphs.
 */
const JP_FULL = '…‥‐←→↑↓―∥、。ー～＝＿￣｜（）〔〕［］｛｝〈〉《》「」『』【】〖〗';
const MEIRYO_YU = '…‥‐←→↑↓―、。ー～＝＿￣｜（）〔〕［］｛｝〈〉《》「」『』【】〖〗';
const SC_SONG = '…‥←→↑↓―∥、。ー～＝＿￣｜（）〔〕［］｛｝〈〉《》「」『』【】〖〗';
const SC_HEI = '…←→↑↓、。～＿￣｜（）〔〕［］｛｝〈〉《》「」『』【】〖〗';
const DENGXIAN = '←→↑↓∥、。ー～＝＿￣｜（）〔〕［］｛｝〈〉《》「」『』【】〖〗';
const YAHEI = '…‥←→↑↓―∥、。ー～＝＿（）〔〕［］｛｝〈〉《》「」『』【】〖〗';
const KOREAN = '…‥、。（）〔〕［］｛｝〈〉《》「」『』【】';
const MINGLIU = '…‥（）〔〕［］｛｝〈〉《》「」『』【】';
const VERTICAL_GLYPHS: ReadonlyMap<string, ReadonlySet<number>> = new Map(
  ([
    [['ms gothic', 'ms pgothic', 'ms ui gothic', 'ms mincho', 'ms pmincho', 'hggothice', 'hgpgothice', 'hgsgothice',
      'hgminchoe', 'hgpminchoe', 'hgsminchoe', 'hgmarugothicmpro', 'hgsoeikakugothicub', 'hgpsoeikakugothicub',
      'hgssoeikakugothicub'], JP_FULL],
    [['yu gothic'], MEIRYO_YU + '∥'],
    [['meiryo', 'meiryo ui', 'yu gothic ui', 'yu mincho', 'yu mincho demibold'], MEIRYO_YU],
    [['yu gothic medium'], '…‥‐←→↑↓―、。ー～＝＿￣｜（）〔〕［］｛｝《》「」『』【】〖〗'],
    [['simsun', 'nsimsun', 'stzhongsong'], SC_SONG],
    [['simhei', 'kaiti', 'fangsong', 'sthupo', 'stliti', 'stxingkai', 'stxinwei'], SC_HEI],
    [['dengxian'], DENGXIAN],
    [['microsoft yahei', 'microsoft yahei ui'], YAHEI],
    [['microsoft jhenghei'], '…‥、。ー（）〔〕［］｛｝〈〉《》「」『』【】〖〗'],
    [['batang', 'batangche', 'gungsuh', 'gungsuhche', 'gulim', 'gulimche', 'dotum', 'dotumche', 'malgun gothic'], KOREAN],
    [['mingliu', 'pmingliu'], MINGLIU],
    [['mingliu_hkscs'], '…‥（）〔〕｛｝〈〉《》「」『』【】'],
  ] as const).flatMap(([families, chars]) => {
    const set = new Set([...chars].map((c) => c.codePointAt(0)!));
    return families.map((f) => [f, set] as const);
  }),
);

/** Whether `family` draws the sideways character `cp` with a vertical glyph. */
export function stackedVerticalGlyph(family: string, cp: number): boolean {
  const set = VERTICAL_GLYPHS.get(canonicalFamily(family)) ?? VERTICAL_GLYPHS.get(normalizeFamily(family));
  return set?.has(cp) ?? false;
}
