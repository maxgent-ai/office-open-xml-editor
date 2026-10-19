import { describe, expect, it } from 'vitest';
import { buildSegments, layoutLines } from './line-layout.js';
import { createLayoutServices } from './layout-runtime.js';
import type { DocRun, DocxDocumentModel } from './types.js';

// A zero-ink custom note-mark host probes its painting face only when its face
// has no line metric (an absent authored face). A present face keeps its
// established metric-only geometry, including authored baseline positions.
const ctx = {
  font: '10px serif', letterSpacing: '0px', fontKerning: 'none',
  measureText(text: string) {
    return { width: text.length * 5, fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2,
      actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2 } as TextMetrics;
  },
} as unknown as CanvasRenderingContext2D;
const model = {
  section: { pageWidth: 612, pageHeight: 792, marginTop: 72, marginRight: 72, marginBottom: 72, marginLeft: 72,
    headerDistance: 36, footerDistance: 36, titlePage: false, evenAndOddHeaders: false },
  body: [], headers: { default: null, first: null, even: null }, footers: { default: null, first: null, even: null },
} as unknown as DocxDocumentModel;

describe('custom note-mark host metric probe', () => {
  it('keeps present custom-note hosts on their metric-only geometry', () => {
    for (const position of [0, 2]) {
      const services = createLayoutServices(model, { measureContext: ctx, localMetrics: { arial: {
        requestedFamily: 'Arial', family: 'Arial', weight: 400, style: 'normal', sourceIdentity: 'office-local:arial',
        lineHeightRatio: 1.15, designAscentRatio: 0.9, designDescentRatio: 0.25,
      } } });
      const run = { type: 'text', text: '', fontFamily: 'Arial', fontSize: 10, bold: false, italic: false, position,
        noteRef: { kind: 'footnote', id: '1', customMarkFollows: true } } as unknown as DocRun;
      const segs = buildSegments([run], { pageIndex: 0, totalPages: 1, layoutServices: services });
      expect(segs.every((segment) => !('metricProbeText' in segment) || segment.metricProbeText === undefined)).toBe(true);
      const withoutProbe = segs.map((segment) => ({ ...segment, metricProbeText: undefined }));
      const geometry = (line: { ascent: number; descent: number }) => ({ ascent: line.ascent, descent: line.descent });
      expect(geometry(layoutLines(ctx, segs, 100, 0, 1)[0]!)).toEqual(geometry(layoutLines(ctx, withoutProbe, 100, 0, 1)[0]!));
    }
  });

  it('does not probe a present face that has no prepared line metric', () => {
    for (const position of [0, 2]) {
      // A positively selected local resource without numeric vertical metrics
      // is still present: its host keeps the established metric-only geometry.
      const services = createLayoutServices(model, { measureContext: ctx, localMetrics: { arial: {
        requestedFamily: 'Arial', family: 'Arial', weight: 400, style: 'normal', sourceIdentity: 'office-local:arial',
      } } });
      const run = { type: 'text', text: '', fontFamily: 'Arial', fontSize: 10, bold: false, italic: false, position,
        noteRef: { kind: 'footnote', id: '1', customMarkFollows: true } } as unknown as DocRun;
      const segs = buildSegments([run], { pageIndex: 0, totalPages: 1, layoutServices: services });
      expect(segs.some((segment) => 'fontSource' in segment && segment.fontSource === 'local')).toBe(true);
      expect(segs.every((segment) => !('metricProbeText' in segment) || segment.metricProbeText === undefined)).toBe(true);
    }
  });

  it('probes the painting face when the authored face is absent', () => {
    const services = createLayoutServices(model, { measureContext: ctx });
    const run = { type: 'text', text: '', fontFamily: 'Absent Authored Face', fontSize: 10, bold: false, italic: false,
      noteRef: { kind: 'footnote', id: '1', customMarkFollows: true } } as unknown as DocRun;
    const segs = buildSegments([run], { pageIndex: 0, totalPages: 1, layoutServices: services });
    expect(segs.some((segment) => 'metricProbeText' in segment && segment.metricProbeText === 'x')).toBe(true);
    const line = layoutLines(ctx, segs, 100, 0, 1)[0]!;
    expect(line.ascent + line.descent).toBeCloseTo(10, 8);
  });
});
