import { describe, expect, it } from 'vitest';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import type { DocParagraph, DocxDocumentModel, DocRun, LineSpacing } from './types.js';
import type { ParagraphLayout } from './layout/types.js';

// Word for Mac printing PDFs, modes 14/15: coordinates are quantized to
// 0.24pt. Expectations below are the measured nominal geometry, not Canvas
// font boxes (the deliberately different stub cannot supply the Word rule).
function render(font: string, size: number, spacing: LineSpacing, grid = false, hostSize?: number, mode = 15, underline = false, markSize = size, wrap = false, localAlias: boolean | 'provided' = false, decoratedLast = false, markCsFace?: string) {
  let canvasFont = '';
  const context = {
    get font() { return canvasFont; }, set font(v: string) { canvasFont = v; },
    measureText(text: string) {
      const px = Number(/([\d.]+)px/.exec(canvasFont)?.[1] ?? 10);
      return { width: text.length * px / 3, fontBoundingBoxAscent: px * .7,
        fontBoundingBoxDescent: px * .3, actualBoundingBoxAscent: px * .7,
        actualBoundingBoxDescent: px * .3 } as TextMetrics;
    },
  } as CanvasRenderingContext2D;
  const text = (value: string): DocRun => ({ type: 'text', text: value,
    fontFamily: font, fontSize: size, bold: false, italic: false,
    underline, strikethrough: false, color: null, isLink: false,
    background: null, vertAlign: null } as DocRun);
  const runs: DocRun[] = [];
  for (let i = 0; i < 3; i++) {
    if (hostSize !== undefined) runs.push({ type: 'anchorHost', fontFamily: font,
      fontFamilyEastAsia: font, fontSize: hostSize, bold: false, italic: false } as DocRun);
    runs.push({ ...text('PROBE Hgyp'), ...(decoratedLast && i === 2 ? { underline: true } : {}) } as DocRun);
    if (i < 2) runs.push({ type: 'break', breakType: 'line', fontSize: size } as DocRun);
  }
  if (wrap) runs.push({ type: 'image', imagePath: 'word/media/test.png',
    mimeType: 'image/png', widthPt: 80, heightPt: 60, anchor: true,
    anchorXPt: 80, anchorYPt: 0, anchorXRelativeFrom: 'margin',
    anchorYRelativeFrom: 'margin', anchorYFromPara: false,
    wrapMode: 'square', wrapSide: 'bothSides' } as DocRun);
  const paragraph = { type: 'paragraph', runs, alignment: 'left', indentLeft: 0,
    indentRight: 0, indentFirst: 0, spaceBefore: 0, spaceAfter: 0,
    lineSpacing: spacing, tabStops: [], numbering: null, defaultFontSize: markSize,
    defaultFontFamily: font, widowControl: false,
    ...(markCsFace ? { paragraphMarkFontFacts: { fontFamily: font, fontFamilyCs: markCsFace,
      fontSize: markSize } } : {}),
  } as DocParagraph;
  const empty = { default: null, first: null, even: null };
  const model = { body: [paragraph], headers: empty, footers: empty, settings: { compatibilityMode: mode },
    section: { pageWidth: 612, pageHeight: 792, marginTop: 72, marginLeft: 72,
      marginRight: 72, marginBottom: 72, headerDistance: 36, footerDistance: 36,
      titlePage: false, evenAndOddHeaders: false,
      ...(grid ? { docGridType: 'lines', docGridLinePitch: 20 } : {}) },
  } as DocxDocumentModel;
  const layout = layoutDocument(model, createLayoutServices(model, {
    measureContext: context,
    ...(localAlias ? { localMetrics: { [font.toLowerCase()]: {
      family: 'Installed test face', requestedFamily: font,
      weight: 400, style: 'normal' as const, sourceIdentity: localAlias === 'provided'
        ? 'provided-sfnt:test-face' : 'office-local:test-installed-face',
      lineHeightRatio: 1.14990234375, designAscentRatio: .9052734375,
      designDescentRatio: .2119140625,
    } } } : {}),
  }));
  return layout.pages[0].layers.body.find(n => n.kind === 'paragraph') as ParagraphLayout;
}

describe('Word specified spacing for ordinary Latin text', () => {
  it.each([
    ['Arial', 10, 30, false], ['Verdana', 10, 24, true],
    ['Yu Mincho', 30, 12, true],
  ] as const)('places %s exact text independently of its font box and anchor hosts', (font, size, height, grid) => {
    const free = render(font, size, { rule: 'exact', value: height, explicit: true }, grid);
    const hosted = render(font, size, { rule: 'exact', value: height, explicit: true }, grid, 30);
    for (const line of free.lines) expect(line.baselinePt - line.bounds.yPt).toBeCloseTo(height * 4 / 5, 10);
    expect(hosted.lines.map(l => l.baselinePt)).toEqual(free.lines.map(l => l.baselinePt));
    expect(hosted.flowBounds.heightPt).toBe(3 * height);
  });

  it('keeps the unmeasured compatibility-mode class on the established centered path', () => {
    const old = render('Arial', 10, { rule: 'exact', value: 30, explicit: true }, false, undefined, 12);
    expect(old.lines[0].baselinePt - 72).toBe(17);
    const fourteen = render('Arial', 10, { rule: 'exact', value: 30, explicit: true }, false, undefined, 14);
    expect(fourteen.lines[0].baselinePt - 72).toBe(24);
  });

  it('preserves exact heights outside the measured evidence interval', () => {
    const paragraph = render('Arial', 10, { rule: 'exact', value: 40, explicit: true });
    expect(paragraph.lines[0].baselinePt - 72).toBe(22);
    expect(paragraph.flowBounds.heightPt).toBe(120);
  });

  it('preserves the placement contract for float-exclusion-conditioned lines', () => {
    const paragraph = render('Arial', 10, { rule: 'exact', value: 30, explicit: true },
      false, undefined, 15, false, 10, true);
    expect(paragraph.lines[0].baselinePt - paragraph.lines[0].bounds.yPt).toBe(17);
  });

  it('keeps installed local aliases in the authored-face baseline class', () => {
    const paragraph = render('Arial', 10, { rule: 'exact', value: 30, explicit: true },
      false, 30, 15, false, 10, false, true);
    expect(paragraph.lines[0].baselinePt - 72).toBe(24);
  });

  it('preserves selected application-provided resources instead of borrowing authored reference geometry', () => {
    const spacing = { rule: 'exact', value: 30, explicit: true } as const;
    const paragraph = render('Arial', 10, spacing, false, undefined, 15, false, 10, false, 'provided');
    const previous = render('Arial', 10, spacing, false, undefined, 12, false, 10, false, 'provided');
    expect(paragraph.lines.map(line => line.baselinePt)).toEqual(previous.lines.map(line => line.baselinePt));
  });

  it('puts non-grid atLeast leading before text without reserving a tall floating host', () => {
    const spacing = { rule: 'atLeast', value: 24, explicit: true } as const;
    const free = render('Arial', 10, spacing);
    const hosted = render('Arial', 10, spacing, false, 30);
    expect(free.lines[0].baselinePt - 72).toBeCloseTo(21.880859375, 10);
    expect(hosted.lines.map(l => l.baselinePt)).toEqual(free.lines.map(l => l.baselinePt));
    expect(hosted.flowBounds.heightPt).toBe(72);
  });

  it('preserves unmeasured decorated text and a differently-sized paragraph mark', () => {
    const spacing = { rule: 'exact', value: 30, explicit: true } as const;
    const underlined = render('Arial', 10, spacing, false, undefined, 15, true);
    const differentMark = render('Arial', 10, spacing, false, undefined, 15, false, 12);
    expect(underlined.lines[0].baselinePt - 72).toBe(17);
    expect(differentMark.lines[0].baselinePt - 72).toBe(17);
  });

  it('preserves the whole paragraph when decoration starts on a later physical line', () => {
    const paragraph = render('Arial', 10, { rule: 'exact', value: 30, explicit: true },
      false, undefined, 15, false, 10, false, false, true);
    expect(paragraph.lines.map(line => line.baselinePt - line.bounds.yPt)).toEqual([17, 17, 17]);
  });

  it('preserves a paragraph mark with a different script-axis face', () => {
    const paragraph = render('Arial', 10, { rule: 'exact', value: 30, explicit: true },
      false, undefined, 15, false, 10, false, false, false, 'Times New Roman');
    expect(paragraph.lines[0].baselinePt - 72).toBe(17);
  });

  it('allocates whole grid cells for the explicit single-line diagnostic', () => {
    const paragraph = render('Yu Mincho', 30, { rule: 'auto', value: 1, explicit: true }, true);
    expect(paragraph.lines[0].advancePt).toBe(60);
    expect(Math.abs(paragraph.lines[0].baselinePt - 72 - 40)).toBeLessThan(.24);
  });

  it('preserves tall-host allocation in the unresolved grid-atLeast class', () => {
    const hosted = render('Arial', 10, { rule: 'atLeast', value: 24, explicit: true }, true, 30);
    expect(hosted.flowBounds.heightPt).toBeCloseTo(103.4912109375, 10);
  });

  it.each([
    ['Arial', 10, 24, 24, 15.63037109375],
    ['Yu Mincho', 30, 24, 42.97998046875001, 31.355712890625],
  ] as const)('preserves the unresolved %s grid-atLeast class', (font, size, minimum, advance, baseline) => {
    const paragraph = render(font, size, { rule: 'atLeast', value: minimum, explicit: true }, true);
    expect(paragraph.lines[0].advancePt).toBe(advance);
    expect(paragraph.lines[0].baselinePt - 72).toBeCloseTo(baseline, 10);
  });
});
