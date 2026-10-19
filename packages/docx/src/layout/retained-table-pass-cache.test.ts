import { describe, expect, it } from 'vitest';
import { createLayoutServices } from '../layout-runtime.js';
import type { DocxDocumentModel } from '../types.js';
import type { BodyAcquisitionState } from './acquisition-context.js';
import { retainedTablePassKey } from './retained-table-pass-cache.js';
import { createParagraphAcquisitionCacheServicesView } from './runtime-state.js';

function model(): DocxDocumentModel {
  return {
    section: {
      pageWidth: 612, pageHeight: 792,
      marginTop: 72, marginRight: 72, marginBottom: 72, marginLeft: 72,
      headerDistance: 36, footerDistance: 36,
      titlePage: false, evenAndOddHeaders: false,
    },
    body: [],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
  } as DocxDocumentModel;
}

function measureContext(): CanvasRenderingContext2D {
  return {
    font: '',
    letterSpacing: '0px',
    fontKerning: 'auto',
    measureText: (text: string) => ({ width: text.length * 5 }) as TextMetrics,
  } as unknown as CanvasRenderingContext2D;
}

function state(overrides: Partial<BodyAcquisitionState> = {}): BodyAcquisitionState {
  const services = createParagraphAcquisitionCacheServicesView(
    createLayoutServices(model(), { measureContext: measureContext() }),
  );
  const sectionLayout = {
    geometry: { pageWidth: 612, pageHeight: 792, marginTop: 72, marginRight: 72, marginBottom: 72, marginLeft: 72 },
    columns: [{ xPt: 72, wPt: 468 }],
    columnSeparator: false,
    grid: { kind: 'none', linePitchPt: null, charSpacePt: null },
    textDirection: 'lrTb',
    sectionBidi: false,
    verticalAlignment: 'top',
  };
  return {
    ctx: {},
    verticalGlyphMeasurement: {},
    acquisitionInputs: {},
    contentX: 72,
    contentW: 468,
    y: 100,
    pageH: 792,
    pageIndex: 0,
    totalPages: 1,
    marginLeft: 72,
    marginRight: 72,
    marginTop: 72,
    marginBottom: 72,
    pageWidth: 612,
    layoutSettings: {},
    sectionLayout,
    storyContext: { story: 'body', containers: [], lineNumberingEligible: true },
    docEastAsian: false,
    fontFamilyClasses: {},
    resolvedLocalFonts: {},
    layoutServices: services,
    retainedTableAcquisition: {},
    retainedTablesBySourceIndex: new Map(),
    kinsoku: {},
    defaultTabPt: 36,
    floats: [],
    floatParaSeq: 0,
    ...overrides,
  } as unknown as BodyAcquisitionState;
}

describe('retained table pass memo key', () => {
  it('is equal for states that differ only in fields table acquisition replaces or rebuilds', () => {
    const base = state();
    const key = retainedTablePassKey(base, 3, 468);
    expect(key).toBeDefined();
    // Cell states replace the cursor, exclusions and prescan set; sessions
    // rebuild their closures and per-session table records identically.
    const other: BodyAcquisitionState = {
      ...base,
      y: 400,
      floats: [{ x: 0, y: 0, w: 10, h: 10 } as unknown as BodyAcquisitionState['floats'][number]],
      floatParaSeq: 7,
      pageAnchorPrescanned: new Set(),
      frozenAnchorFrames: new Map([['a', { xPt: 0, yPt: 0, widthPt: 1, heightPt: 1 }]]),
      retainedTablesBySourceIndex: new Map(),
      retainedTableAcquisition: {} as BodyAcquisitionState['retainedTableAcquisition'],
      // A section context rebuilt for another page with the same content.
      sectionLayout: structuredClone(base.sectionLayout),
    };
    expect(retainedTablePassKey(other, 3, 468)).toBe(key);
  });

  it('differs for any location, page, field or document-object difference', () => {
    const base = state();
    const key = retainedTablePassKey(base, 3, 468);
    expect(retainedTablePassKey(base, 4, 468)).not.toBe(key);
    expect(retainedTablePassKey(base, 3, 400)).not.toBe(key);
    const variants: Partial<BodyAcquisitionState>[] = [
      { contentX: 300 },
      { pageIndex: 1 },
      { totalPages: 2 },
      { displayPageNumber: 5 },
      { pageH: 700 },
      { marginTop: 50 },
      { containerShading: 'FF0000' },
      { uprightPhysicalTable: true },
      { noteReferenceNumber: 2 },
      { currentDateMs: 1 },
      { sectionLayout: { ...base.sectionLayout, textDirection: 'tbRl' } },
      { layoutSettings: {} as BodyAcquisitionState['layoutSettings'] },
      { fontFamilyClasses: {} },
    ];
    for (const variant of variants) {
      expect(retainedTablePassKey({ ...base, ...variant }, 3, 468)).not.toBe(key);
    }
  });

  it('declines states it cannot key exactly', () => {
    expect(retainedTablePassKey(state({ revisionAuthorColor: () => '#000000' }), 3, 468))
      .toBeUndefined();
    const unscoped = state({
      layoutServices: createLayoutServices(model(), { measureContext: measureContext() }),
    });
    expect(retainedTablePassKey(unscoped, 3, 468)).toBeUndefined();
  });
});
