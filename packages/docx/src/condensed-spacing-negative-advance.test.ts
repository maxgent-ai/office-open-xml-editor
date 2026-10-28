import { describe, it, expect } from 'vitest';
import { layoutBodyModel } from './test-support/document-layout.test-support.js';
import { segAdvanceWidth, type LayoutTextSeg } from './line-layout.js';
import type { LineLayout, TextPlacement } from './layout/types.js';
import type { BodyElement, DocParagraph, DocxTextRun, SectionProps } from './types';

// #1975 / WORD_CONDENSED_SPACING_ADVANCE_FLOOR. ECMA-376 §17.3.2.35 removes
// condensed pitch "after each character"; Word floors every character's
// advance at zero instead of moving the pen backwards. A Word export of
// "AB" + "Mi"/"iM" (w:spacing -10 pt, Arial 20 pt) places "CD" at AB + (M - 10)
// in both orders, and an over-condensed run collapses its underline and
// shading to nothing.

/** Stub glyph advance: 'i' is 0.2 em, every other character 0.5 em. */
function makeCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = px();
      return {
        width: [...s].reduce((sum, ch) => sum + p * (ch === 'i' ? 0.2 : 0.5), 0),
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
  };
  return ctx as unknown as CanvasRenderingContext2D;
}

const section: SectionProps = {
  pageWidth: 400, pageHeight: 400,
  marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
  headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
};

function textRun(text: string, extra: Partial<DocxTextRun> & { charSpacing?: number } = {}): DocParagraph['runs'][number] {
  const run: DocxTextRun = {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 20, color: null, fontFamily: 'NotInMetrics', isLink: false, background: null,
    vertAlign: null, hyperlink: null, ...extra,
  };
  return { type: 'text', ...run } as DocParagraph['runs'][number];
}

function para(runs: DocParagraph['runs'], extra: Partial<DocParagraph> = {}): BodyElement {
  const p: DocParagraph = {
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs, defaultFontSize: 20, defaultFontFamily: 'NotInMetrics', ...extra,
  };
  return { type: 'paragraph', ...p } as BodyElement;
}

function lines(body: BodyElement[]): LineLayout[] {
  const node = layoutBodyModel(body, section, makeCtx()).pages[0]!.layers.body[0]!;
  if (node.kind !== 'paragraph') throw new Error('expected a paragraph');
  return [...node.lines];
}

function texts(line: LineLayout): TextPlacement[] {
  return line.placements.filter((p): p is TextPlacement => p.kind === 'text');
}

const byText = (line: LineLayout, text: string): TextPlacement =>
  texts(line).find((p) => p.text === text)!;

describe('#1975 condensed w:spacing larger than the glyph advance', () => {
  it('floors each over-condensed character at zero advance instead of rejecting', () => {
    // M = 10 pt, spacing -30 pt: each M advances 0, so CD starts where AB ends.
    const [line] = lines([para([textRun('AB'), textRun('MM', { charSpacing: -30 }), textRun('CD')])]);
    const run = byText(line!, 'MM');
    expect(run.advancePt).toBe(0);
    expect(run.bounds.widthPt).toBe(0);
    expect(run.clusters.map((c) => [c.offset.xPt, c.advancePt])).toEqual([[0, 0], [0, 0]]);
    expect(byText(line!, 'CD').bounds.xPt).toBe(byText(line!, 'AB').bounds.xPt + 20);
    // Each floored glyph paints from its own retained origin.
    expect(run.paintOps.map((op) => [op.text, op.offset.xPt])).toEqual([['M', 0], ['M', 0]]);
  });

  it.each(['Mi', 'iM'])('floors per character, not per run (%s)', (text) => {
    // M = 10 → 4, i = 4 → -2 floored to 0: the run advances 4 in both orders.
    // A signed or per-run sum would give 2.
    const [line] = lines([para([textRun('AB'), textRun(text, { charSpacing: -6 }), textRun('CD')])]);
    const run = byText(line!, text);
    expect(run.advancePt).toBeCloseTo(4, 9);
    expect(byText(line!, 'CD').bounds.xPt).toBeCloseTo(byText(line!, 'AB').bounds.xPt + 24, 9);
    const offsets = run.clusters.map((c) => [c.offset.xPt, c.advancePt]);
    expect(offsets).toEqual(text === 'Mi' ? [[0, 4], [4, 0]] : [[0, 0], [0, 4]]);
    // The unfloored M keeps native letterSpacing; the floored i is separate.
    expect(run.paintOps.map((op) => [op.text, op.offset.xPt, op.letterSpacingPt]))
      .toEqual(text === 'Mi' ? [['M', 0, -6], ['i', 4, -6]] : [['i', 0, -6], ['M', 0, -6]]);
  });

  it('floors a condensed space run', () => {
    const [line] = lines([para([textRun('Lorem'), textRun(' ', { charSpacing: -30 }), textRun('ipsum')])]);
    expect(byText(line!, ' ').advancePt).toBe(0);
    expect(byText(line!, 'ipsum').bounds.xPt).toBe(byText(line!, 'Lorem').bounds.xPt + 50);
  });

  it('aligns right and centre with the floored width', () => {
    const runs = () => [textRun('AB'), textRun('MMMM', { charSpacing: -30 }), textRun('CD')];
    const [right] = lines([para(runs(), { alignment: 'right' })]);
    expect(byText(right!, 'CD').bounds.xPt + 20).toBeCloseTo(380, 9);
    expect(byText(right!, 'AB').bounds.xPt).toBeCloseTo(340, 9);
    const [center] = lines([para(runs(), { alignment: 'center' })]);
    expect(byText(center!, 'AB').bounds.xPt).toBeCloseTo(180, 9);
  });

  it('breaks lines with the floored width, never borrowing room from negative advances', () => {
    // 360 pt band. "Q" + 20 over-condensed M (0 pt) + 36 'x' (360 pt): a signed
    // run (-400 pt) would pull every x onto line 1; floored, they overflow it.
    const [first, second] = lines([para([
      textRun('Q'), textRun('M'.repeat(20), { charSpacing: -30 }), textRun('x'.repeat(36)),
    ])]);
    expect(second).toBeDefined();
    const firstWidth = texts(first!).reduce((sum, p) => sum + p.advancePt, 0);
    expect(firstWidth).toBeLessThanOrEqual(360 + 1e-9);
    expect(texts(first!).every((p) => p.advancePt >= 0)).toBe(true);
  });

  it('justifies a line holding a floored run without negative geometry', () => {
    const words = 'aa bb cc dd ee ff gg hh ii jj kk ll mm nn oo pp qq rr ss tt uu vv ';
    const [first] = lines([para([
      textRun(words), textRun('MMMM', { charSpacing: -30 }), textRun(words),
    ], { alignment: 'both' })]);
    const placed = texts(first!);
    expect(placed.every((p) => p.bounds.widthPt >= 0)).toBe(true);
    const last = placed.at(-1)!;
    expect(last.bounds.xPt + last.bounds.widthPt).toBeLessThanOrEqual(380 + 1e-6);
  });

  it('collapses the decoration of a fully floored run', () => {
    const [line] = lines([para([
      textRun('AB'), textRun('MMMM', { charSpacing: -30, underline: true }), textRun('CD'),
    ])]);
    const lengths = line!.placements.flatMap((p) => p.kind === 'text' ? p.decorations ?? [] : [])
      .map((d) => 'to' in d && 'from' in d ? (d as { to: { xPt: number }; from: { xPt: number } }).to.xPt
        - (d as { from: { xPt: number } }).from.xPt : 0);
    expect(lengths.every((length) => length === 0)).toBe(true);
  });
});

describe('segAdvanceWidth condensed floor', () => {
  const seg = (text: string, charSpacing: number): LayoutTextSeg => ({
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 20, color: null, fontFamily: 'X', vertAlign: null, measuredWidth: 0, charSpacing,
  } as LayoutTextSeg);

  it('sums per-cluster floors when the caller supplies clusters', () => {
    const clusters = [{ text: 'M', advancePx: 10 }, { text: 'i', advancePx: 4 }];
    expect(segAdvanceWidth(seg('Mi', -6), 14, undefined, 1, clusters)).toBeCloseTo(4, 12);
  });

  it('is unchanged for ordinary and mildly condensed spacing', () => {
    expect(segAdvanceWidth(seg('Mi', -2), 14, undefined, 1)).toBe(10);
    expect(segAdvanceWidth(seg('Mi', 3), 14, undefined, 1)).toBe(20);
  });
});
