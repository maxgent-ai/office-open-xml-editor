import { referenceFontLineMetrics } from '../reference-font-line-metrics.js';
import { layoutDocument } from '../document-layout.js';
import { createLayoutServices } from '../layout-runtime.js';
import { paintLayoutPage } from '../paint/canvas-page.js';
import type { DocxDocumentModel, DocParagraph, DocxTextRun } from '../types.js';
import { makeRecordingCanvas } from './recording-canvas.test-support.js';

export const spacingFixtures = [
  { id: 'tracking', tracking: 0.35, scale: 1 },
  { id: 'scaling', tracking: -0.15, scale: 0.66 },
  { id: 'grid', tracking: 0.35, scale: 0.66, grid: 1200 },
  { id: 'review-justified', tracking: 0.35, scale: 0.66, grid: 1200, alignment: 'both' },
  { id: 'prefix-counts', tracking: 0.35, scale: 0.66, grid: 1200, text: 'AVToWAVEAVATAR'.repeat(4) },
  { id: 'vertical', tracking: 0.35, scale: 0.66, grid: 1200, vertical: true },
  { id: 'vertical-upright', tracking: 0.35, scale: 0.66, grid: 1200, vertical: true, text: '甲乙（AVATAR）😀甲乙'.repeat(3) },
  { id: 'kashida', tracking: 0.35, scale: 0.66, grid: 1200, alignment: 'highKashida', arabic: true },
] as const;

export async function recordSpacingFixture(
  fixture: { id: string; tracking: number; scale: number; grid?: number;
    alignment?: string; vertical?: boolean; arabic?: boolean; text?: string },
  nativeCanvas?: HTMLCanvasElement,
) {
  const run: DocxTextRun & { type: 'text' } = {
    type: 'text', text: fixture.text ?? (fixture.arabic ? 'بببب بببب ' + 'W'.repeat(60) : 'AV To WAVE AVATAR '.repeat(4).trimEnd()),
    fontFamily: 'Arial', fontFamilyCs: 'Arial', fontSize: 20,
    bold: false, italic: false, underline: false, strikethrough: false,
    color: null, isLink: false, background: null, vertAlign: null, hyperlink: null,
    kerning: 20, charSpacing: fixture.tracking, charScale: fixture.scale,
    ...(fixture.arabic ? { cs: true } : {}),
  };
  const model: DocxDocumentModel = {
    section: { pageWidth: 260, pageHeight: 600, marginTop: 0, marginRight: 0,
      marginBottom: 0, marginLeft: 0, headerDistance: 0, footerDistance: 0,
      ...(fixture.grid ? { docGridType: 'linesAndChars', docGridCharSpace: fixture.grid,
        docGridLinePitch: 24 } : {}),
      ...(fixture.vertical ? { textDirection: 'tbRl' } : {}),
    },
    body: [{ type: 'paragraph', alignment: fixture.alignment ?? 'left',
      indentLeft: 0, indentRight: 0, indentFirst: 0, spaceBefore: 0, spaceAfter: 0,
      lineSpacing: null, numbering: null, tabStops: [], runs: [run],
      defaultFontSize: 20, defaultFontFamily: 'Arial', widowControl: false,
    } as DocParagraph],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
  } as DocxDocumentModel;
  const recording = makeRecordingCanvas({ contextualShaping: true, spaceAdvancePx: 5.5625, advanceFactor: 0.5137 });
  const canvas = nativeCanvas ?? recording.canvas;
  const ctx = canvas.getContext('2d') as CanvasRenderingContext2D;
  // Record native transform arguments too: Firefox's matrix representation can
  // hide or magnify a one-ULP change in an argument to translate().
  const transforms: (string | number)[][] = [];
  for (const name of ['translate', 'scale', 'rotate', 'transform', 'setTransform'] as const) {
    const original = ctx[name]?.bind(ctx) as ((...args: number[]) => void) | undefined;
    if (original) Object.assign(ctx, { [name]: (...args: number[]) => {
      transforms.push([name, ...args]); original(...args);
    } });
  }
  const nativeFills: unknown[][] = [];
  if (nativeCanvas) {
    const original = ctx.fillText.bind(ctx);
    ctx.fillText = (text, x, y) => {
      const matrix = ctx.getTransform();
      nativeFills.push([text, x, y, ctx.font, ctx.letterSpacing, ctx.fontKerning,
        matrix.a, matrix.b, matrix.c, matrix.d, matrix.e, matrix.f]);
      original(text, x, y);
    };
  }
  // Arial is present (its resource metric is the catalog line projection), as
  // in the captured main geometry; an absent face would use its fallback box.
  const arial = { ...referenceFontLineMetrics('Arial', 400, 'normal')!, requestedFamily: 'Arial', family: 'Arial',
    weight: 400, style: 'normal' as const, sourceIdentity: 'local:arial' };
  const layout = layoutDocument(model, createLayoutServices(model, { measureContext: ctx, localMetrics: { arial } }),
    { currentDateMs: 0 });
  const advances = layout.pages.flatMap(page => page.layers.body.flatMap(node =>
    node.kind === 'paragraph' ? node.lines.map(line => line.placements.flatMap(placement =>
      placement.kind === 'text' ? [[placement.text, placement.advancePt,
        placement.clusters.map(cluster => [cluster.offset.xPt, cluster.advancePt])]] : [])) : []));
  for (let page = 0; page < layout.pages.length; page += 1) {
    await paintLayoutPage(layout, page, canvas, { scale: 1, dpr: 1 });
  }
  const fills = nativeCanvas ? nativeFills : recording.fills.map(fill =>
    [fill.text, fill.x, fill.y, fill.font, fill.letterSpacing, fill.fontKerning,
      fill.translateX, fill.translateY, fill.scaleX, fill.scaleY]);
  // Intern repeated font strings without rounding or omitting numeric fields.
  const fonts: string[] = [];
  for (const fill of fills) {
    const font = fill[3] as string;
    if (!fonts.includes(font)) fonts.push(font);
    fill[3] = fonts.indexOf(font);
  }
  return { advances, transforms, fonts, fills };
}
