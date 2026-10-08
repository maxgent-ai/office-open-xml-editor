import { describe, expect, it, vi } from 'vitest';
import { createPageLayers } from './page-graph.js';
import { gapDocument, gapParagraph } from '../test-support/gap-paragraph.test-support.js';
import { fragmentLineAdvancesPt } from './flow-fragment.js';
import { selectParagraphFragment } from './paragraph-pagination.js';
import { composeCanonicalSectionFlow } from './section-flow-composition.js';
import { sliceParagraphLayout, translateParagraphLayout } from './paragraph.js';
import { footnoteIdsInRetainedSlice } from './note-reference-ownership.js';
import { textRunsForPage } from '../text-run-projection.js';
import { DocxFindController } from '../find.js';
import { buildDocxTextLayer } from '../text-layer.js';
import { hitTestDocxElementContext } from '../element-context.js';
import { createPaintResourceRegistry } from './paint-resources.js';
import { paintParagraphLayout } from '../paint/canvas-text.js';
import { exactRetainedColumnBalanceTarget } from './column-balance-frontier.js';
import { pageOwnedAnchorKeysByLine } from './anchor-line-deferral.js';
import type { BodyLayoutSession } from './body-layout-kernel.js';
import type { BodyLayoutInput } from './body-layout-input.js';
import type { CanvasPaintContext } from '../paint/types.js';

function select(available: number, policy = { keepLines: false, widowControl: false }) {
  const { paragraph, boundaries } = gapParagraph();
  return selectParagraphFragment(paragraph, { boundary: null },
    { kind: 'splittable', lineEndBoundaries: boundaries }, available, 100, true, policy);
}

describe('physical lines are the retained consumer boundary', () => {
  it('retains one source range and vertical charge per baseline', () => {
    const { paragraph, boundaries } = gapParagraph();
    expect(paragraph.lines.map(line => [line.range.start, line.range.end, line.placements.length])).toEqual([
      [0, 10, 2], [10, 20, 2], [20, 30, 2],
    ]);
    expect(boundaries).toHaveLength(3);
    expect(fragmentLineAdvancesPt(paragraph)).toBe(30);
    expect(sliceParagraphLayout(paragraph, { lineStart: 0, lineEnd: 1,
      continuesFromPrevious: false, continuesOnNext: true }).advancePt).toBe(10);
  });

  it('numbers each baseline once and advances the section counter by physical lines', () => {
    const { paragraph } = gapParagraph();
    const session = { measureLineNumberGlyph: () => ({ widthPt: 5, ascentPt: 8, descentPt: 2 }) } as unknown as BodyLayoutSession;
    const next = { ...gapParagraph(1).paragraph, id: 'next-paragraph' };
    const document = gapDocument(paragraph);
    const page = document.pages[0]!;
    const input = { ...document, pages: [{ ...page, layers: createPageLayers([
      { layer: 'body', node: paragraph }, { layer: 'body', node: next },
    ]) }] };
    const result = composeCanonicalSectionFlow(input, session, []);
    const retained = result.pages[0]!.layers.body[0]!;
    if (retained.kind !== 'paragraph') throw new Error('missing paragraph');
    expect(retained.lineNumbers?.map(number => [number.counterValue, number.paintOps[0]?.origin.yPt])).toEqual([
      [1, paragraph.lines[0]!.baselinePt], [2, paragraph.lines[1]!.baselinePt], [3, paragraph.lines[2]!.baselinePt],
    ]);
    const following = result.pages[0]!.layers.body[1]!;
    if (following.kind !== 'paragraph') throw new Error('missing next paragraph');
    expect(following.lineNumbers?.map(number => number.counterValue)).toEqual([4]);
  });

  it('moves three physical lines instead of leaving a one-line widow', () => {
    expect(select(20, { keepLines: false, widowControl: true }).requiresFreshFlowRegion).toBe(true);
  });

  it('keeps all gaps when keepLines moves the paragraph', () => {
    expect(select(10, { keepLines: true, widowControl: false }).requiresFreshFlowRegion).toBe(true);
  });

  it('owns every gap on the same page and resumes after the complete baseline', () => {
    const result = select(10);
    expect(result.fragment?.lines).toHaveLength(1);
    expect(result.fragment?.lines[0]?.placements).toHaveLength(2);
    expect(result.nextCursor?.boundary).toEqual({ segIndex: 2, charOffset: 0 });
  });

  it('keeps note references in later gaps with the reference page', () => {
    const { paragraph } = gapParagraph();
    const slice = (start: number, end: number) => sliceParagraphLayout(paragraph, {
      lineStart: start, lineEnd: end, continuesFromPrevious: start > 0, continuesOnNext: end < 3,
    });
    expect(footnoteIdsInRetainedSlice(slice(0, 1))).toEqual(['1']);
    expect(footnoteIdsInRetainedSlice(slice(1, 3))).toEqual(['3']);
  });

  it('assigns a later-gap page anchor to its physical line deferral slot', () => {
    const { paragraph } = gapParagraph();
    const owner = paragraph.lines.findIndex(line => line.placements.some(placement =>
      placement.kind === 'text' && placement.text === 'BBBB '));
    const bounds = { xPt: 40, yPt: 0, widthPt: 120, heightPt: 100 };
    const anchored = { ...paragraph, drawings: [{
      kind: 'drawing' as const, id: 'anchor', source: paragraph.source,
      flowDomainId: 'body', ordinaryFlow: false, flowBounds: bounds, inkBounds: bounds,
      advancePt: 0, commands: [], anchorLayer: { occurrenceId: 'page-anchor',
        horizontalOwnership: 'page' as const, verticalOwnership: 'page' as const,
        behindDoc: false, relativeHeight: 0, sourceOrder: 0 },
    }], lines: paragraph.lines.map((line, index) => index !== owner ? line : {
      ...line, placements: [...line.placements, { kind: 'drawing' as const, range: { start: 5, end: 5 },
        drawingId: 'anchor', bounds, advancePt: 0 }],
    }) };
    expect(pageOwnedAnchorKeysByLine(anchored)).toEqual([['page-anchor'], [], []]);
  });

  it('balances widow-controlled lines without treating gaps as extra lines', () => {
    const { paragraph } = gapParagraph(2);
    const document = gapDocument(paragraph);
    const page = document.pages[0]!;
    const region = { ...page.sectionRegions[0]!, flowDomainIds: ['body', 'second-column'] };
    const input = { initialSection: { sectionOccurrenceId: 'section' }, sequence: [{ kind: 'body-block', block: { kind: 'paragraph', source: paragraph.source,
      keepLines: false, keepNext: false, widowControl: true } }] } as unknown as BodyLayoutInput;
    expect(exactRetainedColumnBalanceTarget(input, [{ nodeId: paragraph.id, flowDomainId: 'body',
      blockStartPt: 0, blockEndPt: 20 }], new Map(), page, region)).toBe(20);
  });

  it('projects selection and accessibility runs in source order without spanning the obstacle', () => {
    const runs = textRunsForPage(gapDocument(gapParagraph(1).paragraph), 0, { scale: 1 });
    expect(runs.map(run => [run.text, run.x, run.w])).toEqual([['AAAA ', 0, 25], ['BBBB ', 160, 25]]);
    const element = () => ({ style: {}, dataset: {}, children: [] as unknown[], appendChild(child: unknown) { this.children.push(child); } });
    const layer = { ...element(), ownerDocument: { createElement: element }, innerHTML: '' };
    buildDocxTextLayer(layer as unknown as HTMLDivElement, runs, 200, 800);
    expect(layer.children).toHaveLength(2);
    expect(layer.children.map(child => (child as { textContent: string }).textContent)).toEqual(['AAAA ', 'BBBB ']);
    expect(runs[0]!.y).toBe(runs[1]!.y);
  });

  it('finds text across gap seams with two disjoint highlight slices', async () => {
    const runs = textRunsForPage(gapDocument(gapParagraph(1).paragraph), 0, { scale: 1 });
    const find = new DocxFindController(() => 1, async () => runs);
    expect(await find.find('AAAA BBBB')).toHaveLength(1);
    expect(find.pageHighlights(0)[0]?.slices).toHaveLength(2);
  });

  it('hit-tests a picture in the later gap and rejects the excluded area', () => {
    const p = gapParagraph(1, false, true).paragraph;
    const resource = p.lines[0]!.placements.find(placement => placement.kind === 'resource');
    if (resource?.kind !== 'resource') throw new Error('missing picture');
    const registry = createPaintResourceRegistry([{ kind: 'image', resourceKey: resource.resourceKey,
      partPath: 'word/media/picture.png', mimeType: 'image/png', intrinsicSize: { widthPt: 20, heightPt: 10 } }]);
    const document = gapDocument(p);
    expect(hitTestDocxElementContext(document, 0, { xPt: 165, yPt: resource.bounds.yPt + 5 }, registry)).toMatchObject({
      elementType: 'image', source: { path: [0, 1] }, bounds: { xPt: 160, widthPt: 20 },
    });
    expect(hitTestDocxElementContext(document, 0, { xPt: 100, yPt: resource.bounds.yPt + 5 }, registry)).toBeNull();
  });

  it('does not justify either gap of the paragraph-final physical line', () => {
    const p = gapParagraph(1, true).paragraph;
    expect(p.lines[0]!.placements.map(placement => 'advancePt' in placement ? placement.advancePt : 0)).toEqual([25, 25]);
  });

  it('translates and shades the horizontal allocations without covering the float', () => {
    const { paragraph } = gapParagraph(1);
    const translated = translateParagraphLayout(paragraph, { xPt: 7, yPt: 11 });
    expect(translated.lines[0]!.wrapFragments?.map(box => [box.xPt, box.yPt, box.widthPt])).toEqual([[7, 11, 40], [167, 11, 40]]);
    const fillRect = vi.fn();
    paintParagraphLayout({ ...translated, shading: { color: '#abcdef' }, lines: translated.lines.map(line => ({ ...line, placements: [] })) },
      { ctx: { fillRect }, viewport: { scale: 1 }, pageTransform: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 } } as unknown as CanvasPaintContext);
    expect(fillRect.mock.calls).toEqual([[7, 11, 40, 10], [167, 11, 40, 10]]);
  });
});
