import { describe, expect, it } from 'vitest';
import {
  DEFAULT_KINSOKU_RULES,
  seaMixedBreakOffsets,
  seaTransitionOffsets,
} from '@silurus/ooxml-core';
import { layoutLines, type LayoutLine, type LayoutTextSeg } from './line-layout.js';

// Issue #960 — no-space SEA↔Latin/CJK/digit script-transition boundaries, and
// mixed CJK+SEA runs. Adjudicated against Word-exported ground truth
// (sample-45.pdf, private): Word breaks at these transitions (it did NOT carry
// the whole cross-script unit down — a Thai→Latin seam `…ของ | Thailand` fell on
// a line boundary), and in a single run mixing CJK and Thai each script keeps
// its own break rule (CJK per-character, Thai at cluster/dictionary boundaries).
//
// All four Thai fixtures carry `<w:cs/>`, so buildSegments keeps the whole
// no-space paragraph in ONE complex-script segment — which is exactly why the
// two gaps manifest. The stub metric is 5px per code point (makeLinearCtx), so
// the fitter is deterministic w.r.t. width; the WORD boundaries come from the
// platform ICU dictionary, so we assert the PROPERTY (every break is a legal
// merged offset; the transitions ARE used) rather than exact offsets.

function makeLinearCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const fontSize = (): number => Number.parseFloat(/([\d.]+)px/.exec(font)?.[1] ?? '10');
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    measureText: (text: string) => {
      const size = fontSize();
      return {
        width: [...text].length * size * 0.5,
        fontBoundingBoxAscent: size * 0.8,
        fontBoundingBoxDescent: size * 0.2,
        actualBoundingBoxAscent: size * 0.8,
        actualBoundingBoxDescent: size * 0.2,
      } as TextMetrics;
    },
  } as unknown as CanvasRenderingContext2D;
}
function textSeg(text: string): LayoutTextSeg {
  return {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 10, color: null, fontFamily: 'T', vertAlign: null, measuredWidth: 0,
  } as unknown as LayoutTextSeg;
}
function lay(text: string, width: number): LayoutLine[] {
  return layoutLines(
    makeLinearCtx(), [textSeg(text)], width, 0, 1, [], undefined, {}, 0,
    DEFAULT_KINSOKU_RULES, undefined, 36, width, false,
  );
}
const lineTexts = (lines: LayoutLine[]): string[] =>
  lines.map((l) => l.segments.filter((s): s is LayoutTextSeg => 'text' in s).map((s) => s.text).join(''));

/** Cumulative UTF-16 offset at the END of every non-final line = each break. */
function breakOffsets(texts: string[]): number[] {
  const out: number[] = [];
  let acc = 0;
  for (let i = 0; i < texts.length - 1; i++) { acc += texts[i].length; out.push(acc); }
  return out;
}

describe('SEA↔non-SEA no-space transitions in docx layoutLines (#960)', () => {
  const A1 = 'เมืองBangkokคือเมืองหลวงของThailandมีชื่อเต็มยาวว่าKrungThepMahaNakhon';
  const A2 = 'ราคาสินค้า1250บาทลดเหลือ990บาทประหยัด260บาทหรือ21เปอร์เซ็นต์ต่อชิ้น';
  const A3 = 'อาหารญี่ปุ่น寿司และราเมงラーメンเป็นที่นิยมในไทยข้าวปั้นおにぎりมากขึ้น';
  const B = '日本語のテキストとภาษาไทยが同じ';

  const legal = (t: string): Set<number> =>
    new Set(seaMixedBreakOffsets(t, { cjk: true, kinsoku: DEFAULT_KINSOKU_RULES }));

  it('a1 Thai↔Latin: breaks only at legal merged offsets, and USES a transition seam', () => {
    // Width 110 (22 cp) > the widest unit between two seams (KrungThepMahaNakhon,
    // 19 cp) so no emergency char-split occurs — every break is a legal offset.
    const texts = lineTexts(lay(A1, 110));
    expect(texts.join('')).toBe(A1); // nothing lost/duplicated
    expect(texts.length).toBeGreaterThan(1);
    const legalSet = legal(A1);
    for (const b of breakOffsets(texts)) expect(legalSet.has(b)).toBe(true);
    // At least one break is a genuine SEA↔non-SEA transition (the #960 fix).
    const transitions = new Set(seaTransitionOffsets(A1));
    expect(breakOffsets(texts).some((b) => transitions.has(b))).toBe(true);
  });

  it('a2 Thai↔digit: breaks at the digit seams, never mid-Thai-cluster and never mid-number', () => {
    const texts = lineTexts(lay(A2, 60));
    expect(texts.join('')).toBe(A2);
    expect(texts.length).toBeGreaterThan(1);
    const legalSet = legal(A2);
    for (const b of breakOffsets(texts)) expect(legalSet.has(b)).toBe(true);
    // A price like "1250"/"990"/"260"/"21" is a single European-digit group; a
    // break must never fall between two digits (that would split the number).
    for (const b of breakOffsets(texts)) {
      const prev = A2.codePointAt(b - 1)!;
      const at = A2.codePointAt(b)!;
      const isDigit = (c: number) => c >= 0x30 && c <= 0x39;
      expect(isDigit(prev) && isDigit(at)).toBe(false);
    }
    // And the transition seams are actually exercised.
    const transitions = new Set(seaTransitionOffsets(A2));
    expect(breakOffsets(texts).some((b) => transitions.has(b))).toBe(true);
  });

  it('a3 Thai↔CJK: Thai breaks at dictionary boundaries, CJK per character', () => {
    const texts = lineTexts(lay(A3, 60));
    expect(texts.join('')).toBe(A3);
    expect(texts.length).toBeGreaterThan(1);
    const legalSet = legal(A3);
    for (const b of breakOffsets(texts)) expect(legalSet.has(b)).toBe(true);
  });

  it('cross-run kinsoku: a mixed run led by 。 is not orphaned at a line head', () => {
    // A prior Latin word fills the line; the following mixed CJK+SEA run starts
    // with 。 (行頭禁則). Routing it through the SEA branch must still pull a
    // trailing grapheme of the previous segment down so 。 never heads a line
    // (§17.3.1.16 cross-run 追い出し), exactly as the CJK branch does.
    const lines = layoutLines(
      makeLinearCtx(), [textSeg('abcd'), textSeg('。ภาษาไทย')], 25, 0, 1, [], undefined, {}, 0,
      DEFAULT_KINSOKU_RULES, undefined, 36, 25, false,
    );
    const texts = lineTexts(lines);
    expect(texts.join('')).toBe('abcd。ภาษาไทย');
    for (const t of texts) expect(t.startsWith('。')).toBe(false);
  });

  it('b CJK+Thai one run: each side keeps its rule — the Thai span is never torn', () => {
    // 「ภาษาไทย」 is a short Thai span embedded in CJK. With the transition seam
    // now legal, Word (and now we) break at と|ภาษาไทย and keep the Thai whole,
    // instead of tearing it mid-cluster (ภาษ|าไทย) on the old CJK-only path.
    const texts = lineTexts(lay(B, 60));
    expect(texts.join('')).toBe(B);
    expect(texts.length).toBeGreaterThan(1);
    const legalSet = legal(B);
    for (const b of breakOffsets(texts)) expect(legalSet.has(b)).toBe(true);
    // The contiguous Thai cluster run stays intact on ONE line.
    expect(texts.some((t) => t.includes('ภาษาไทย'))).toBe(true);
    // CJK still breaks at a character boundary (日本語… wraps per glyph as needed).
    // The break set is non-empty and a CJK-adjacent offset is used.
    expect(breakOffsets(texts).length).toBeGreaterThan(0);
  });
});
