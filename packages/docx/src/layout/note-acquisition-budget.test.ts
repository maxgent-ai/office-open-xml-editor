import { describe, expect, it } from 'vitest';
import { createLayoutServices } from '../layout-runtime.js';
import { layoutDocument } from '../document-layout.js';
import { layoutSourceStore } from '../layout-source-model-adapter.js';
import { noteContinuationModel } from '../testing/note-continuation-model.js';
import { layoutDocumentInputAsync } from './document.js';
import type { DocParagraph } from '../types.js';
import { referenceFontLineMetrics } from '../reference-font-line-metrics.js';

/** The fixture's Times New Roman is present here, as in the convergence
 * geometry these budgets were sized for (its resource metric is the catalog
 * projection). An absent face would size lines from its painting fallback. */
function presentTimes() {
  return { 'times new roman': {
    ...referenceFontLineMetrics('Times New Roman', 400, 'normal')!,
    requestedFamily: 'Times New Roman', family: 'Times New Roman', weight: 400, style: 'normal' as const,
    sourceIdentity: 'local:times-new-roman',
  } };
}

function measureContext(): CanvasRenderingContext2D {
  return { font: '10px serif', letterSpacing: '0px', fontKerning: 'auto',
    measureText: (text: string) => ({ width: [...text].length * 5,
      actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2,
      fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2 } as TextMetrics),
  } as unknown as CanvasRenderingContext2D;
}

function largeSourceNote(paragraphs: number) {
  const model = noteContinuationModel();
  model.section.pageHeight = 60;
  const content = model.footnotes![0]!.content;
  model.footnotes![0]!.content = Array.from({ length: paragraphs }, (_, index) => {
    const original = content[index % content.length] as DocParagraph;
    return { ...original, type: 'paragraph', runs: [...original.runs] };
  });
  return model;
}

describe('cumulative full-footnote acquisition work', () => {
  it('rejects repeated page-dependent acquisition across convergence passes', async () => {
    const model = largeSourceNote(30);
    const paragraph = model.footnotes![0]!.content[0] as DocParagraph;
    const run = paragraph.runs[0];
    if (run?.type !== 'text') throw new Error('Expected text fixture');
    // Metadata costs source traversal/snapshot work even though it paints only
    // one PAGE result. Keep geometry small while exercising a near-limit source.
    paragraph.runs.push({ ...run, type: 'field', fieldType: 'page',
      instruction: 'PAGE' + ' '.repeat(900_000), fallbackText: '1' });
    const services = createLayoutServices(model, {
      measureContext: measureContext(), allowFootnoteContinuation: true, localMetrics: presentTimes(),
    });
    const source = layoutSourceStore(model);
    const progress: number[] = [];
    await expect(layoutDocumentInputAsync(source.bodyLayoutInput, services,
      { currentDateMs: 0 }, { sliceMs: 0, yieldToHost: async () => undefined,
        onProgress: pages => progress.push(pages) }))
      .rejects.toThrow('Footnote acquisition cumulative work budget exceeded');
    // The first pass fits below the work ceiling; a new field service view
    // must inherit its debit instead of obtaining another complete allowance.
    expect(progress.some((pages, index) => index > 0 && pages < progress[index - 1]!)).toBe(true);
  }, 20_000);

  it('does not charge a reused immutable whole note on each destination page', () => {
    const model = largeSourceNote(80);
    const paragraph = model.footnotes![0]!.content[0] as DocParagraph;
    paragraph.styleId = 'S'.repeat(900_000);
    const services = createLayoutServices(model, {
      measureContext: measureContext(), allowFootnoteContinuation: true, localMetrics: presentTimes(),
    });
    const layout = layoutDocument(model, services, { currentDateMs: 0 });
    expect(layout.pages.length).toBeGreaterThan(32);
    expect(layout.pages.flatMap(page => page.layers.notes).flatMap(note =>
      note.kind === 'note' ? note.story.blocks : [])).toHaveLength(80);
  }, 20_000);

  it('bounds unbounded endnote acquisition even when footnote continuation is disabled', () => {
    const model = largeSourceNote(1);
    const note = model.footnotes![0]!;
    (note.content[0] as DocParagraph).styleId = 'S'.repeat(1_000_001);
    for (const block of model.body) {
      if (block.type !== 'paragraph') continue;
      for (const run of block.runs) if (run.type === 'text' && run.noteRef) run.noteRef.kind = 'endnote';
    }
    model.endnotes = [note];
    model.footnotes = [];
    const services = createLayoutServices(model, {
      measureContext: measureContext(), allowFootnoteContinuation: false, localMetrics: presentTimes(),
    });
    expect(() => layoutDocument(model, services, { currentDateMs: 0 }))
      .toThrow('Footnote acquisition source budget exceeded');
  });
});
