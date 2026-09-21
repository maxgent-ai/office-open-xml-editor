import { afterEach, describe, expect, it } from 'vitest';
import { DEFAULT_KINSOKU_RULES, parseOpenTypeLineMetrics, unregisterEmbeddedFonts } from '@silurus/ooxml-core';
import { measureParagraph } from './paragraph-measure.js';
import { createLayoutServices } from './layout-runtime.js';
import { buildSegments, docGridLineCells, layoutLines, paragraphMarkLineMetrics } from './line-layout.js';
import type { DocParagraph, DocRun, DocxDocumentModel } from './types.js';
import {
  loadDocxFontResources,
  snapshotDocxFontResources,
} from './font-resources.js';

const globals = globalThis as Record<string, unknown>;
const original = { document: globals.document, self: globals.self, FontFace: globals.FontFace };

afterEach(() => {
  globals.document = original.document;
  globals.self = original.self;
  globals.FontFace = original.FontFace;
});

function installFontFaceSet(): void {
  class FakeFontFace {
    family: string;
    weight: string;
    style: string;
    status = 'loaded';
    constructor(
      family: string,
      _source: ArrayBuffer,
      descriptors?: { weight?: string; style?: string },
    ) {
      this.family = family;
      this.weight = descriptors?.weight ?? 'normal';
      this.style = descriptors?.style ?? 'normal';
    }
    load(): Promise<FakeFontFace> { return Promise.resolve(this); }
  }
  globals.FontFace = FakeFontFace;
  globals.document = { fonts: { add() {}, delete() {}, ready: Promise.resolve() } };
  delete globals.self;
}

function metricSfnt(first = 0x41, last = 0x41): Uint8Array {
  const tableCount = 4;
  const headOffset = 12 + tableCount * 16;
  const hheaOffset = headOffset + 54;
  const os2Offset = hheaOffset + 36;
  const cmapOffset = os2Offset + 78;
  const bytes = new Uint8Array(cmapOffset + 40);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, 0x00010000);
  view.setUint16(4, tableCount);
  const record = (index: number, tag: string, offset: number, length: number) => {
    const at = 12 + index * 16;
    for (let i = 0; i < 4; i++) bytes[at + i] = tag.charCodeAt(i);
    view.setUint32(at + 8, offset);
    view.setUint32(at + 12, length);
  };
  record(0, 'head', headOffset, 54);
  record(1, 'hhea', hheaOffset, 36);
  record(2, 'OS/2', os2Offset, 78);
  record(3, 'cmap', cmapOffset, 40);
  view.setUint16(headOffset + 18, 1000);
  view.setInt16(hheaOffset + 4, 800);
  view.setInt16(hheaOffset + 6, -200);
  view.setInt16(hheaOffset + 8, 100);
  view.setUint16(os2Offset + 62, 0x0080);
  view.setInt16(os2Offset + 68, 700);
  view.setInt16(os2Offset + 70, -180);
  view.setInt16(os2Offset + 72, 20);
  view.setUint16(cmapOffset + 2, 1);
  view.setUint16(cmapOffset + 4, 3);
  view.setUint16(cmapOffset + 6, 10);
  view.setUint32(cmapOffset + 8, 12);
  view.setUint16(cmapOffset + 12, 12);
  view.setUint32(cmapOffset + 16, 28);
  view.setUint32(cmapOffset + 24, 1);
  view.setUint32(cmapOffset + 28, first);
  view.setUint32(cmapOffset + 32, last);
  view.setUint32(cmapOffset + 36, 1);
  return bytes;
}

describe('application-provided DOCX font resources', () => {
  it('owns the input bytes and derives tuple metrics from the supplied resource', async () => {
    installFontFaceSet();
    const source = metricSfnt();
    const snapshot = snapshotDocxFontResources([{
      family: 'Arbitrary Resource Face',
      bytes: source,
      weight: 700,
      style: 'italic',
    }]);
    source.fill(0);

    const first = await loadDocxFontResources(snapshot);
    const second = await loadDocxFontResources(snapshot);
    const metric = first.metrics['arbitrary resource face:700:italic'];
    expect(metric).toMatchObject({
      requestedFamily: 'Arbitrary Resource Face',
      weight: 700,
      style: 'italic',
      lineHeightRatio: 0.9,
      designAscentRatio: 0.7,
      designDescentRatio: 0.18,
      lineGapRatio: 0.02,
      unicodeRanges: [[0x41, 0x41]],
    });
    expect(metric.family).toMatch(/^__ooxml_provided_[0-9a-f]{64}$/);
    expect(second.metrics['arbitrary resource face:700:italic'].family).toBe(metric.family);
    unregisterEmbeddedFonts(first.faces);
    unregisterEmbeddedFonts(second.faces);
  });

  it('bounds admitted faces before copying them across the worker boundary', () => {
    const bytes = metricSfnt();
    const resources = Array.from({ length: 65 }, (_, index) => ({ family: `Face ${index}`, bytes }));
    expect(snapshotDocxFontResources(resources)).toHaveLength(64);
  });

  it('does not turn an unreadable font or an empty family into layout metrics', async () => {
    installFontFaceSet();
    const resources = snapshotDocxFontResources([
      { family: '', bytes: metricSfnt() },
      { family: 'Broken', bytes: new Uint8Array([0, 1, 2, 3]) },
    ]);
    await expect(loadDocxFontResources(resources)).resolves.toMatchObject({ metrics: {} });
  });
});

// Independent geometry fixture: glyphs advance 10pt, loose Canvas box 10pt,
// actual ink 7pt. Supplied sfnt owns a 7pt ascent + 1.8pt descent + .2pt gap.
function layoutFixture(metrics: Awaited<ReturnType<typeof loadDocxFontResources>>['metrics'] = {}, embedded = false) {
  const context = {
    font: '10px serif', letterSpacing: '0px', fontKerning: 'auto',
    measureText: (text: string) => ({
      width: [...text].length * 10,
      actualBoundingBoxAscent: 6, actualBoundingBoxDescent: 1,
      fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2,
    }),
  } as unknown as CanvasRenderingContext2D;
  const model = {
    section: { pageWidth: 612, pageHeight: 792, marginTop: 72, marginRight: 72,
      marginBottom: 72, marginLeft: 72, headerDistance: 36, footerDistance: 36,
      titlePage: false, evenAndOddHeaders: false }, body: [],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
  } as DocxDocumentModel;
  if (embedded) model.embeddedFonts = [{ fontName: 'Arbitrary Resource Face',
    style: 'regular', partPath: 'word/fonts/font1.odttf', fontKey: '' }];
  const services = createLayoutServices(model, { measureContext: context, localMetrics: metrics,
    embeddedFaces: embedded ? [{ family: 'Arbitrary Resource Face', weight: 'normal',
      style: 'normal', status: 'loaded' } as FontFace] : [],
  });
  return {
    services, context,
    lines(text: string, width = 1000, family = 'Arbitrary Resource Face', bold = false, size = 10) {
      const runs = [{ type: 'text', text, fontFamily: family, fontFamilyEastAsia: family, fontSize: size,
        bold, italic: false, underline: false, strikethrough: false,
      }] as DocRun[];
      const segments = buildSegments(runs, { pageIndex: 0, totalPages: 1, layoutServices: services });
      return layoutLines(context, segments, width, 0, 1);
    },
  };
}

describe('opt-in resource layout boundary', () => {
  it('keeps legacy parsing and default wrapping while exact resources use natural fit', async () => {
    installFontFaceSet();
    const bytes = metricSfnt(0x20, 0x7e);
    // Ordinary embedded consumers retain the pre-feature result shape.
    expect(parseOpenTypeLineMetrics(bytes)).not.toHaveProperty('unicodeRanges');
    const absent = layoutFixture();
    const empty = layoutFixture((await loadDocxFontResources([])).metrics);
    const loaded = await loadDocxFontResources(snapshotDocxFontResources([
      { family: 'Arbitrary Resource Face', bytes },
    ]));
    try {
      const exact = layoutFixture(loaded.metrics);
      expect(absent.lines('AAAA BBBB CCCC', 138)).toHaveLength(1);
      expect(empty.lines('AAAA BBBB CCCC', 138)).toHaveLength(1);
      expect(exact.lines('AAAA BBBB CCCC', 138)).toHaveLength(2);
      const baselineLine = absent.lines('A')[0]!;
      const exactLine = exact.lines('A')[0]!;
      expect(baselineLine.ascent + baselineLine.descent).toBe(10);
      expect(exactLine.ascent).toBeCloseTo(7);
      expect(exactLine.descent).toBeCloseTo(1.8);
      const mark = { runs: [], defaultFontFamily: 'Arbitrary Resource Face',
        defaultFontSize: 10, lineSpacing: null } as unknown as DocParagraph;
      const markMetrics = (fixture: ReturnType<typeof layoutFixture>) => paragraphMarkLineMetrics(
        mark, 1, undefined, false, false, fixture.context, {}, null,
        fixture.services.text.fontMetrics, fixture.services.text,
      );
      expect(markMetrics(exact).advancePx).toBeCloseTo(9);
      expect(markMetrics(absent).advancePx).toBe(10);
      // An authored embedded font keeps the pre-existing path for the same tuple.
      const embedded = layoutFixture(loaded.metrics, true);
      expect(embedded.lines('A')[0]!.ascent + embedded.lines('A')[0]!.descent).toBe(10);
      expect(markMetrics(embedded).advancePx).toBe(10);
      // A different family and a missing weight tuple keep their baseline box.
      for (const line of [exact.lines('A', 1000, 'Other Face')[0]!,
        exact.lines('A', 1000, 'Arbitrary Resource Face', true)[0]!]) {
        expect(line.ascent + line.descent).toBe(10);
      }
    } finally { unregisterEmbeddedFonts(loaded.faces); }
  });

  it('keeps the unknown East-Asian grid count unless a covered resource supplies its box', async () => {
    installFontFaceSet();
    const bytes = metricSfnt(0x7269, 0x7269);
    const hheaOffset = 12 + 4 * 16 + 54;
    const view = new DataView(bytes.buffer);
    view.setInt16(hheaOffset + 4, 500);
    view.setInt16(hheaOffset + 6, -100);
    const loaded = await loadDocxFontResources(snapshotDocxFontResources([
      { family: 'Arbitrary Resource Face', bytes },
    ]));
    try {
      const baseline = layoutFixture().lines('物', 1000, 'Arbitrary Resource Face', false, 20)[0]!;
      const fixture = layoutFixture(loaded.metrics);
      const covered = fixture.lines('物', 1000, 'Arbitrary Resource Face', false, 20)[0]!;
      const uncovered = fixture.lines('理', 1000, 'Arbitrary Resource Face', false, 20)[0]!;
      expect(baseline.gridCountSingle).toBeCloseTo(26);
      expect(uncovered.gridCountSingle).toBeCloseTo(26);
      expect(covered.gridCountSingle).toBeCloseTo(20 * .6 * 1.3);
      expect(docGridLineCells(baseline.gridCountSingle!, 20)).toBe(2);
      expect(docGridLineCells(covered.gridCountSingle!, 20)).toBe(1);
    } finally { unregisterEmbeddedFonts(loaded.faces); }
  });

  it('uses the exact mark design line for inline-image leading despite a legacy family profile', async () => {
    installFontFaceSet();
    const loaded = await loadDocxFontResources(snapshotDocxFontResources([
      { family: 'Calibri', bytes: metricSfnt(0x20, 0x7e) },
    ]));
    try {
      const fixture = layoutFixture(loaded.metrics);
      const auto = { rule: 'auto' as const, value: 1.15, explicit: true };
      const paragraph = { defaultFontFamily: 'Calibri', defaultFontSize: 10,
        lineSpacing: auto, spaceBefore: 0, spaceAfter: 0,
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
      }, { startYPt: 0, paragraphXPt: 0, availableWidthPt: 200, maximumYPt: 300,
        suppressSpaceBefore: false,
      }, { context: fixture.context, fontFamilyClasses: {} }, {
        pageIndex: 0, totalPages: 1, pageWritingMode: 'horizontal-tb',
        documentHasEastAsianText: false, layoutServices: fixture.services,
        resolvedLocalFonts: fixture.services.text.fontMetrics,
      });
      expect(measured.lines[0]?.advancePt).toBeCloseTo(50 + .15 * 9);
    } finally { unregisterEmbeddedFonts(loaded.faces); }
  });

  it('does not lend subset metrics or natural-fit authority to uncovered text', async () => {
    installFontFaceSet();
    const loaded = await loadDocxFontResources(snapshotDocxFontResources([
      { family: 'Arbitrary Resource Face', bytes: metricSfnt() },
    ]));
    try {
      const fixture = layoutFixture(loaded.metrics);
      const uncovered = fixture.lines('AB')[0]!;
      expect(uncovered.ascent + uncovered.descent).toBe(10);
      expect(fixture.lines('AAAA BBBB CCCC', 138)).toHaveLength(1);
    } finally { unregisterEmbeddedFonts(loaded.faces); }
  });
});


describe('exact resource baseline geometry', () => {
  it('keeps the resource ascent and descent when its design box exceeds the Canvas box', async () => {
    installFontFaceSet();
    const bytes = metricSfnt(0x20, 0x7e);
    const os2Offset = 12 + 4 * 16 + 54 + 36;
    const view = new DataView(bytes.buffer);
    view.setInt16(os2Offset + 68, 1000);
    view.setInt16(os2Offset + 70, -200);
    view.setInt16(os2Offset + 72, 0);
    const loaded = await loadDocxFontResources(snapshotDocxFontResources([
      { family: 'Arbitrary Resource Face', bytes },
    ]));
    try {
      const line = layoutFixture(loaded.metrics).lines('A')[0]!;
      expect(line.ascent).toBeCloseTo(10);
      expect(line.descent).toBeCloseTo(2);
    } finally { unregisterEmbeddedFonts(loaded.faces); }
  });

  it('selects the acquired paragraph-mark tuple rather than the paragraph default', async () => {
    installFontFaceSet();
    const loaded = await loadDocxFontResources(snapshotDocxFontResources([
      { family: 'Arbitrary Resource Face', bytes: metricSfnt(0x20, 0x7e) },
    ]));
    try {
      const fixture = layoutFixture(loaded.metrics);
      const mark = { runs: [], defaultFontFamily: 'Other Face', defaultFontSize: 10,
        lineSpacing: null } as unknown as DocParagraph;
      const result = paragraphMarkLineMetrics(mark, 1, undefined, false, false,
        fixture.context, {}, null, fixture.services.text.fontMetrics, fixture.services.text, {
          fonts: { ascii: 'Arbitrary Resource Face', highAnsi: 'Arbitrary Resource Face' },
          fontSizePt: 10, weight: 400, style: 'normal', complexScript: false,
        });
      expect(result.advancePx).toBeCloseTo(9);
      expect(result.ascentPx).toBeCloseTo(7);
    } finally { unregisterEmbeddedFonts(loaded.faces); }
  });

  it('uses the selected resource design baseline for a contentless floating-anchor host', async () => {
    installFontFaceSet();
    const loaded = await loadDocxFontResources(snapshotDocxFontResources([
      { family: 'Arbitrary Resource Face', bytes: metricSfnt(0x20, 0x7e) },
    ]));
    try {
      const fixture = layoutFixture(loaded.metrics);
      const segments = buildSegments([
        { type: 'anchorHost', fontSize: 10, fontFamily: 'Arbitrary Resource Face' },
      ], {
        pageIndex: 0, totalPages: 1, layoutServices: fixture.services,
        resolvedLocalFonts: fixture.services.text.fontMetrics,
      });
      const line = layoutLines(fixture.context, segments, 1000, 0, 1)[0]!;
      expect(line.ascent).toBeCloseTo(7);
      expect(line.descent).toBeCloseTo(1.8);
    } finally { unregisterEmbeddedFonts(loaded.faces); }
  });
});
