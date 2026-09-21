import { describe, expect, it } from 'vitest';
import { DEFAULT_KINSOKU_RULES, type ResolvedFontMetric } from '@silurus/ooxml-core';
import { measureParagraph } from './paragraph-measure.js';
import { createLayoutServices } from './layout-runtime.js';
import { buildSegments, docGridLineCells, layoutLines, paragraphMarkLineMetrics } from './line-layout.js';
import type { DocParagraph, DocRun, DocxDocumentModel } from './types.js';

const FAMILY = 'Arbitrary Resource Face';

function metric(
  family = FAMILY,
  unicodeRanges: readonly (readonly [number, number])[] = [[0x20, 0x7e]],
  overrides: Partial<ResolvedFontMetric> = {},
): ResolvedFontMetric {
  return Object.freeze({
    family: `__test_resolved_${family.toLowerCase().replaceAll(' ', '_')}`,
    requestedFamily: family,
    weight: 400,
    style: 'normal',
    // Resource-derived metrics are the exact-authority class under test. The
    // bytes-to-metric parser is covered separately in core; this fixture enters
    // at the production DOCX layout-service boundary.
    sourceIdentity: `provided-sfnt:test:${family.toLowerCase()}`,
    synthesized: false,
    lineHeightRatio: 0.9,
    designAscentRatio: 0.7,
    designDescentRatio: 0.18,
    lineGapRatio: 0.02,
    unicodeRanges,
    ...overrides,
  });
}

function layoutFixture(
  resolved?: ResolvedFontMetric,
  embedded = false,
) {
  const context = {
    font: '10px serif',
    letterSpacing: '0px',
    fontKerning: 'auto',
    measureText: (text: string) => ({
      width: [...text].length * 10,
      actualBoundingBoxAscent: 6,
      actualBoundingBoxDescent: 1,
      fontBoundingBoxAscent: 8,
      fontBoundingBoxDescent: 2,
    }),
  } as unknown as CanvasRenderingContext2D;
  const model = {
    section: {
      pageWidth: 612, pageHeight: 792, marginTop: 72, marginRight: 72,
      marginBottom: 72, marginLeft: 72, headerDistance: 36, footerDistance: 36,
      titlePage: false, evenAndOddHeaders: false,
    },
    body: [],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    ...(embedded ? { embeddedFonts: [{ fontName: FAMILY, style: 'regular',
      partPath: 'word/fonts/font1.odttf', fontKey: '' }] } : {}),
  } as DocxDocumentModel;
  const localMetrics = resolved
    ? { [(resolved.requestedFamily ?? FAMILY).toLowerCase()]: resolved }
    : {};
  const services = createLayoutServices(model, {
    measureContext: context,
    localMetrics,
    embeddedFaces: embedded
      ? [{ family: FAMILY, weight: 'normal', style: 'normal', status: 'loaded' } as FontFace]
      : [],
  });
  return {
    services,
    context,
    lines(text: string, width = 1000, family = FAMILY, bold = false, size = 10) {
      const runs = [{
        type: 'text', text, fontFamily: family, fontFamilyEastAsia: family,
        fontSize: size, bold, italic: false, underline: false, strikethrough: false,
      }] as DocRun[];
      const segments = buildSegments(runs, {
        pageIndex: 0, totalPages: 1, layoutServices: services,
      });
      return layoutLines(context, segments, width, 0, 1);
    },
  };
}

function markMetrics(fixture: ReturnType<typeof layoutFixture>, family = FAMILY) {
  const mark = {
    runs: [], defaultFontFamily: family, defaultFontSize: 10, lineSpacing: null,
  } as unknown as DocParagraph;
  return paragraphMarkLineMetrics(
    mark, 1, undefined, false, false, fixture.context, {}, null,
    fixture.services.text.fontMetrics, fixture.services.text,
  );
}

describe('resolved font metrics layout boundary', () => {
  it('changes only the covered tuple and yields to an authored embedded face', () => {
    const baseline = layoutFixture();
    const exact = layoutFixture(metric());
    expect(baseline.lines('AAAA BBBB CCCC', 138)).toHaveLength(1);
    expect(exact.lines('AAAA BBBB CCCC', 138)).toHaveLength(2);
    expect(baseline.lines('A')[0]!.ascent + baseline.lines('A')[0]!.descent).toBe(10);
    expect(exact.lines('A')[0]!.ascent).toBeCloseTo(7);
    expect(exact.lines('A')[0]!.descent).toBeCloseTo(1.8);
    expect(markMetrics(exact).advancePx).toBeCloseTo(9);
    expect(markMetrics(baseline).advancePx).toBe(10);

    const embedded = layoutFixture(metric(), true);
    expect(embedded.lines('A')[0]!.ascent + embedded.lines('A')[0]!.descent).toBe(10);
    expect(markMetrics(embedded).advancePx).toBe(10);
    expect(exact.lines('A', 1000, 'Other Face')[0]!.ascent + exact.lines('A', 1000, 'Other Face')[0]!.descent).toBe(10);
    expect(exact.lines('A', 1000, FAMILY, true)[0]!.ascent + exact.lines('A', 1000, FAMILY, true)[0]!.descent).toBe(10);
  });

  it('uses a covered East-Asian box for grid cells without lending it to adjacent text', () => {
    const baseline = layoutFixture().lines('物', 1000, FAMILY, false, 20)[0]!;
    const fixture = layoutFixture(metric(FAMILY, [[0x7269, 0x7269]], {
      eastAsianLineHeightRatio: 0.6 * 1.3,
    }));
    const covered = fixture.lines('物', 1000, FAMILY, false, 20)[0]!;
    const uncovered = fixture.lines('理', 1000, FAMILY, false, 20)[0]!;
    expect(baseline.gridCountSingle).toBeCloseTo(26);
    expect(uncovered.gridCountSingle).toBeCloseTo(26);
    expect(covered.gridCountSingle).toBeCloseTo(20 * 0.6 * 1.3);
    expect(docGridLineCells(baseline.gridCountSingle!, 20)).toBe(2);
    expect(docGridLineCells(covered.gridCountSingle!, 20)).toBe(1);
  });

  it('uses the resolved design line for inline-image leading despite a legacy family profile', () => {
    const fixture = layoutFixture(metric('Calibri'));
    const auto = { rule: 'auto' as const, value: 1.15, explicit: true };
    const paragraph = {
      defaultFontFamily: 'Calibri', defaultFontSize: 10, lineSpacing: auto,
      spaceBefore: 0, spaceAfter: 0,
      runs: [{ type: 'image', imagePath: 'word/media/inline.png', mimeType: 'image/png',
        widthPt: 50, heightPt: 50, anchor: false }],
    } as DocParagraph;
    const measured = measureParagraph(paragraph, {
      lineGrid: { active: false, pitchPt: null },
      characterGrid: { active: false, kind: null, pitchPt: null, deltaPt: 0 },
      rightIndentGrid: { pitchPt: null, paragraphAllowsAdjustment: true },
      physicalIndentLeftPt: 0, physicalIndentRightPt: 0, firstIndentPt: 0,
      lineSpacing: auto, spaceBeforePt: 0, spaceAfterPt: 0, baseRtl: false,
      isJustified: false, stretchLastLine: false, tabStops: [], hasRuby: false,
      hasEastAsianText: false, kinsoku: DEFAULT_KINSOKU_RULES, defaultTabPt: 36,
    }, {
      startYPt: 0, paragraphXPt: 0, availableWidthPt: 200, maximumYPt: 300,
      suppressSpaceBefore: false,
    }, { context: fixture.context, fontFamilyClasses: {} }, {
      pageIndex: 0, totalPages: 1, pageWritingMode: 'horizontal-tb',
      documentHasEastAsianText: false, layoutServices: fixture.services,
      resolvedLocalFonts: fixture.services.text.fontMetrics,
    });
    expect(measured.lines[0]?.advancePt).toBeCloseTo(50 + 0.15 * 9);
  });

  it('does not lend subset metrics or natural-fit authority to uncovered text', () => {
    const fixture = layoutFixture(metric(FAMILY, [[0x41, 0x41]]));
    const uncovered = fixture.lines('AB')[0]!;
    expect(uncovered.ascent + uncovered.descent).toBe(10);
    expect(fixture.lines('AAAA BBBB CCCC', 138)).toHaveLength(1);
  });

  it('retains a resolved design box that exceeds the Canvas box', () => {
    const line = layoutFixture(metric(FAMILY, [[0x20, 0x7e]], {
      lineHeightRatio: 1.2, designAscentRatio: 1, designDescentRatio: 0.2,
      lineGapRatio: 0,
    })).lines('A')[0]!;
    expect(line.ascent).toBeCloseTo(10);
    expect(line.descent).toBeCloseTo(2);
  });

  it('selects the acquired paragraph-mark tuple rather than the paragraph default', () => {
    const fixture = layoutFixture(metric());
    const mark = {
      runs: [], defaultFontFamily: 'Other Face', defaultFontSize: 10, lineSpacing: null,
    } as unknown as DocParagraph;
    const result = paragraphMarkLineMetrics(
      mark, 1, undefined, false, false, fixture.context, {}, null,
      fixture.services.text.fontMetrics, fixture.services.text, {
        fonts: { ascii: FAMILY, highAnsi: FAMILY },
        fontSizePt: 10, weight: 400, style: 'normal', complexScript: false,
      },
    );
    expect(result.advancePx).toBeCloseTo(9);
    expect(result.ascentPx).toBeCloseTo(7);
  });

  it('uses the selected design baseline for a contentless floating-anchor host', () => {
    const fixture = layoutFixture(metric());
    const segments = buildSegments([
      { type: 'anchorHost', fontSize: 10, fontFamily: FAMILY },
    ], {
      pageIndex: 0, totalPages: 1, layoutServices: fixture.services,
      resolvedLocalFonts: fixture.services.text.fontMetrics,
    });
    const line = layoutLines(fixture.context, segments, 1000, 0, 1)[0]!;
    expect(line.ascent).toBeCloseTo(7);
    expect(line.descent).toBeCloseTo(1.8);
  });
});
