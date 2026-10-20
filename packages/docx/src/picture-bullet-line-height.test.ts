import { describe, expect, it } from 'vitest';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import { testFontSnapshot } from './layout/test-font-snapshot.js';
import type {
  BodyElement,
  DocxDocumentModel,
  LineSpacing,
  NumberingInfo,
  SectionProps,
} from './types';

// Observed Windows Word (16.0) picture-bullet line heights. A picture bullet
// stands on its paragraph's first baseline and takes part in that line like an
// inline picture of its display height, except under exact spacing. These
// tests pin the rule on a synthetic untabled face (ascent 8, descent 2 at
// 10 pt, natural single 10) so each expected value is plain arithmetic.

const TEST_FONT = 'Synthetic Untabled Serif';

function measureContext(): CanvasRenderingContext2D {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    measureText: (text: string) => {
      const size = px();
      return {
        width: [...text].length * size,
        fontBoundingBoxAscent: size * 0.8,
        fontBoundingBoxDescent: size * 0.2,
        actualBoundingBoxAscent: size * 0.8,
        actualBoundingBoxDescent: size * 0.2,
      } as TextMetrics;
    },
  } as unknown as CanvasRenderingContext2D;
}

function bullet(heightPt: number): NumberingInfo {
  return {
    numId: 1, level: 0, format: 'bullet', text: '', indentLeft: 36, tab: 18, suff: 'tab',
    picBulletImagePath: 'word/media/bullet.png', picBulletMimeType: 'image/png',
    picBulletWidthPt: heightPt, picBulletHeightPt: heightPt,
    fontFacts: { fontSize: 10 },
  } as NumberingInfo;
}

function paragraph(
  text: string,
  lineSpacing: LineSpacing | null,
  numbering: NumberingInfo | null = null,
): BodyElement {
  return {
    type: 'paragraph', alignment: 'left',
    indentLeft: numbering ? 36 : 0, indentRight: 0, indentFirst: numbering ? -18 : 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing, numbering, tabStops: [],
    runs: text === '' ? [] : [{
      type: 'text', text, bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: 10, color: null, fontFamily: TEST_FONT, fontFamilyEastAsia: '',
      isLink: false, background: null, vertAlign: null, hyperlink: null,
    }],
    defaultFontSize: 10, defaultFontFamily: TEST_FONT,
    widowControl: false,
  } as unknown as BodyElement;
}

const auto = (value: number): LineSpacing => ({ value, rule: 'auto', explicit: true });
const atLeast = (value: number): LineSpacing => ({ value, rule: 'atLeast', explicit: true });
const exact = (value: number): LineSpacing => ({ value, rule: 'exact', explicit: true });

/** Lay out a bullet paragraph followed by a plain one; report the bullet box
 * and the bullet paragraph's first baseline and line advance. */
function layout(
  item: BodyElement,
  options: { grid?: number } = {},
) {
  const model = {
    section: {
      pageWidth: 400, pageHeight: 4000,
      marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
      ...(options.grid ? { docGridType: 'lines', docGridLinePitch: options.grid } : {}),
    } as SectionProps,
    body: [item, paragraph('Following', null)],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { [TEST_FONT]: 'roman' },
  } as unknown as DocxDocumentModel;
  const services = createLayoutServices(model, {
    localMetrics: testFontSnapshot([{ family: TEST_FONT }]),
    measureContext: measureContext(),
  });
  const [first] = layoutDocument(model, services).pages[0]!.layers.body;
  if (first?.kind !== 'paragraph') throw new Error('expected a paragraph');
  const picture = first.lines[0]!.placements.find((placement) =>
    placement.kind === 'resource' && placement.resourceKind === 'picture-bullet');
  if (!picture?.bounds) throw new Error('expected a picture-bullet placement');
  return {
    bullet: picture.bounds,
    baselinePt: first.lines[0]!.baselinePt,
    advancePt: first.lines[0]!.advancePt,
  };
}

describe('picture-bullet line height', () => {
  it('raises a single line to the bullet height above the baseline', () => {
    const { bullet: box, baselinePt, advancePt } = layout(paragraph('Item', null, bullet(30)));
    expect(box.yPt + box.heightPt).toBeCloseTo(baselinePt, 6);
    expect(baselinePt).toBeCloseTo(30, 6);
    // 30 pt ascent + 2 pt text descent.
    expect(advancePt).toBeCloseTo(32, 6);
  });

  it('keeps the text line when the bullet is shorter than the text ascent', () => {
    const { baselinePt, advancePt } = layout(paragraph('Item', null, bullet(4)));
    expect(baselinePt).toBeCloseTo(8, 6);
    expect(advancePt).toBeCloseTo(10, 6);
  });

  it('adds only the text leading under an automatic multiple', () => {
    // natural 32 + single 10 x (1.5 - 1).
    const { baselinePt, advancePt } = layout(paragraph('Item', auto(1.5), bullet(30)));
    expect(baselinePt).toBeCloseTo(30, 6);
    expect(advancePt).toBeCloseTo(37, 6);
  });

  it('keeps the larger box under atLeast', () => {
    expect(layout(paragraph('Item', atLeast(18), bullet(30))).advancePt).toBeCloseTo(32, 6);
    expect(layout(paragraph('Item', atLeast(40), bullet(30))).advancePt).toBeCloseTo(40, 6);
  });

  it('ignores the bullet under exact spacing', () => {
    const tall = layout(paragraph('Item', exact(18), bullet(30)));
    const short = layout(paragraph('Item', exact(18), bullet(1)));
    expect(tall.advancePt).toBeCloseTo(18, 6);
    expect(tall.baselinePt).toBeCloseTo(short.baselinePt, 6);
    // The bullet still stands on the baseline and overflows the line's top.
    expect(tall.bullet.yPt + tall.bullet.heightPt).toBeCloseTo(tall.baselinePt, 6);
    expect(tall.bullet.yPt).toBeLessThan(0);
  });

  it('counts whole grid cells from the bullet', () => {
    // 30 pt ascent + 2 pt descent takes two 18 pt cells.
    expect(layout(paragraph('Item', null, bullet(30)), { grid: 18 }).advancePt)
      .toBeCloseTo(36, 6);
  });

  it('counts East Asian grid cells from the bullet and the text below the baseline', () => {
    // A 17 pt bullet alone fits one 18 pt cell; with the East Asian text's
    // descent below the baseline the line takes two.
    expect(layout(paragraph('日本語', null, bullet(17)), { grid: 18 }).advancePt)
      .toBeCloseTo(36, 6);
    expect(layout(paragraph('日本語', null, bullet(4)), { grid: 18 }).advancePt)
      .toBeCloseTo(18, 6);
  });

  it('sizes an empty item from the bullet and pins it under an automatic multiple', () => {
    const single = layout(paragraph('', null, bullet(30)));
    expect(single.advancePt).toBeCloseTo(32, 6);
    expect(single.bullet.yPt).toBeCloseTo(0, 6);
    const double = layout(paragraph('', auto(2), bullet(30)));
    // The whole mark box doubles; the bullet stays at the line's top.
    expect(double.advancePt).toBeCloseTo(64, 6);
    expect(double.bullet.yPt).toBeCloseTo(0, 6);
    expect(double.bullet.yPt + double.bullet.heightPt).toBeCloseTo(double.baselinePt, 6);
  });

  it('keeps an empty item\'s box under exact spacing', () => {
    expect(layout(paragraph('', exact(18), bullet(30))).advancePt).toBeCloseTo(18, 6);
  });

  it('keeps every word when the bullet paragraph continues on the next page', () => {
    const words = Array.from({ length: 30 }, (_, index) => `w${index}`);
    const item = paragraph(words.join(' '), null, bullet(30));
    const model = {
      section: {
        pageWidth: 200, pageHeight: 60,
        marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
        headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
      } as SectionProps,
      body: [item],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      fontFamilyClasses: { [TEST_FONT]: 'roman' },
    } as unknown as DocxDocumentModel;
    const services = createLayoutServices(model, {
      localMetrics: testFontSnapshot([{ family: TEST_FONT }]),
      measureContext: measureContext(),
    });
    const pages = layoutDocument(model, services).pages;
    expect(pages.length).toBeGreaterThan(1);
    const ranges = pages.flatMap((page) => page.layers.body.flatMap((node) =>
      node.kind === 'paragraph' ? node.lines.map((line) => line.range) : []));
    // The lines cover the whole paragraph text without a gap or overlap.
    const text = words.join(' ');
    expect(ranges[0]!.start).toBe(0);
    for (let index = 1; index < ranges.length; index += 1) {
      expect(ranges[index]!.start).toBe(ranges[index - 1]!.end);
    }
    expect(ranges.at(-1)!.end).toBe(text.length);
  });

  it('measures a bullet without an explicit size at the height it paints', () => {
    const numbering = { ...bullet(30), fontFacts: { fontSize: 40 } } as NumberingInfo;
    delete (numbering as { picBulletWidthPt?: number }).picBulletWidthPt;
    delete (numbering as { picBulletHeightPt?: number }).picBulletHeightPt;
    const { bullet: box, baselinePt } = layout(paragraph('Item', null, numbering));
    expect(box.heightPt).toBeCloseTo(40, 6);
    expect(baselinePt).toBeCloseTo(40, 6);
    expect(box.yPt).toBeCloseTo(0, 6);
  });

  it.each([[1, 10], [30, 32]])('keeps a leading break line\'s font box beside a %s pt bullet', (height, advance) => {
    const item = paragraph('Item', null, bullet(height)) as BodyElement & { runs: unknown[] };
    item.runs.unshift({ type: 'break', breakType: 'line' });
    const { advancePt, baselinePt } = layout(item);
    expect(advancePt).toBeCloseTo(advance, 6);
    expect(baselinePt).toBeCloseTo(advance - 2, 6);
  });

  it.each([[undefined, 36], [1, 36], [30, 46.8]] as const)(
    'keeps a leading break\'s own leading under an automatic multiple (bullet %s pt)',
    (height, advance) => {
      const item = paragraph('Item', auto(1.5), height === undefined ? null : bullet(height)) as
        BodyElement & { runs: { fontSize?: number }[] };
      // 24 pt text and break, 10 pt paragraph mark.
      item.runs[0]!.fontSize = 24;
      item.runs.unshift({ type: 'break', breakType: 'line' } as never);
      (item as { paragraphMarkFontFacts?: unknown }).paragraphMarkFontFacts = { fontSize: 10 };
      if (height === undefined) {
        expect(firstAdvance(item)).toBeCloseTo(advance, 6);
      } else {
        expect(layout(item).advancePt).toBeCloseTo(advance, 6);
      }
    },
  );

  it.each([[1, 15], [30, 37]] as const)(
    'takes the text line\'s own leading beside a larger paragraph mark (bullet %s pt)',
    (height, advance) => {
      const item = paragraph('Item', auto(1.5), bullet(height));
      (item as { paragraphMarkFontFacts?: unknown }).paragraphMarkFontFacts = { fontSize: 40 };
      expect(layout(item).advancePt).toBeCloseTo(advance, 6);
    },
  );
});

function firstAdvance(item: BodyElement): number {
  const model = {
    section: {
      pageWidth: 400, pageHeight: 4000,
      marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps,
    body: [item],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { [TEST_FONT]: 'roman' },
  } as unknown as DocxDocumentModel;
  const services = createLayoutServices(model, {
    localMetrics: testFontSnapshot([{ family: TEST_FONT }]),
    measureContext: measureContext(),
  });
  const [first] = layoutDocument(model, services).pages[0]!.layers.body;
  if (first?.kind !== 'paragraph') throw new Error('expected a paragraph');
  return first.lines[0]!.advancePt;
}
