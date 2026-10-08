import { translateBodyOccurrence } from './occurrence-projection.js';
import { replacePageLayerNodes } from './page-graph.js';
import type { BodyLayoutSession } from './body-layout-kernel.js';
import type {
  DocumentLayout,
  LayoutPage,
  LineNumberLayout,
  PaintNode,
  ParagraphLayout,
} from './types.js';
import { wordLineNumberDistancePt } from './section-compatibility.js';

export interface BodyFlowAllocation {
  readonly nodeId: string;
  readonly flowDomainId: string;
  readonly blockStartPt: number;
  readonly blockEndPt: number;
}

function numberedParagraph(
  paragraph: ParagraphLayout,
  counterStart: number,
  countBy: number,
  distancePt: number,
  session: BodyLayoutSession,
): Readonly<{ paragraph: ParagraphLayout; counterEnd: number }> {
  let counter = counterStart;
  // §17.6.8 counts physical lines. Acquisition has already united their gap
  // placements, so the same retained entry owns exactly one counter/baseline.
  const lineNumbers = paragraph.lines.map((line, lineIndex): LineNumberLayout => {
    const counterValue = counter++;
    const text = String(counterValue);
    const metrics = session.measureLineNumberGlyph(text);
    const origin = Object.freeze({
      xPt: paragraph.flowBounds.xPt - distancePt,
      yPt: line.baselinePt,
    });
    return Object.freeze({
      lineIndex,
      counterValue,
      bounds: Object.freeze({
        xPt: origin.xPt - metrics.widthPt,
        yPt: origin.yPt - metrics.ascentPt,
        widthPt: metrics.widthPt,
        heightPt: metrics.ascentPt + metrics.descentPt,
      }),
      paintOps: counterValue % countBy === 0
        ? Object.freeze([Object.freeze({
            kind: 'text' as const,
            text,
            origin,
            font: metrics.font ?? '',
            color: '#000000',
            textAlign: 'right' as const,
          })])
        : Object.freeze([]),
    });
  });
  return Object.freeze({
    paragraph: Object.freeze({ ...paragraph, lineNumbers: Object.freeze(lineNumbers) }),
    counterEnd: counter,
  });
}

/** §17.6.23 and §17.6.8 are section-band composition rules. Admission first
 * fixes page/region ownership; this pure pass then translates retained geometry
 * and attaches counters without measuring or repaginating body content. */
export function composeCanonicalSectionFlow(
  layout: DocumentLayout,
  session: BodyLayoutSession,
  allocations: readonly BodyFlowAllocation[],
): DocumentLayout {
  const sectionCounters = new Map<string, number>();
  let continuousCounter: number | undefined;
  const pages = layout.pages.map((page): LayoutPage => {
    if (page.parityBlank) return page;
    let body = [...page.layers.body];
    for (let regionIndex = 0; regionIndex < page.sectionRegions.length; regionIndex += 1) {
      const region = page.sectionRegions[regionIndex]!;
      const domains = new Set(region.flowDomainIds);
      const indices = body.flatMap((node, index) => domains.has(node.flowDomainId) ? [index] : []);
      const nodesById = new Map(indices.map((index) => [body[index]!.id, body[index]!]));
      const ownedAllocations = allocations.filter((allocation) => {
        const node = nodesById.get(allocation.nodeId);
        return domains.has(allocation.flowDomainId)
          && allocation.blockEndPt > allocation.blockStartPt
          && node !== undefined
          && (node.ordinaryFlow || node.sectionFlowOwnership === 'host-flow');
      });
      const flowTopPt = ownedAllocations.length === 0
        ? region.blockStartPt
        : Math.min(...ownedAllocations.map((allocation) => allocation.blockStartPt));
      const flowBottomPt = ownedAllocations.length === 0
        ? flowTopPt
        : Math.max(...ownedAllocations.map((allocation) => allocation.blockEndPt));
      // A nextColumn boundary can retain two section regions over the same
      // block interval while assigning disjoint physical column domains. Each
      // section's §17.6.23 alignment therefore uses its own retained band, not
      // the next region's block origin.
      const bandBottomPt = region.blockEndPt;
      const bandHeightPt = Math.max(0, bandBottomPt - region.blockStartPt);
      const bodyHeightPt = Math.max(0, flowBottomPt - flowTopPt);
      const alignment = region.section.verticalAlignment;
      const translationPt = ownedAllocations.length > 0 && bodyHeightPt < bandHeightPt
        ? alignment === 'center'
          ? region.blockStartPt + (bandHeightPt - bodyHeightPt) / 2 - flowTopPt
          : alignment === 'bottom'
            ? bandBottomPt - bodyHeightPt - flowTopPt
            : 0
        : 0;

      const numbering = region.section.lineNumbering;
      let counter = numbering?.restart === 'newPage'
        ? numbering.start
        : numbering?.restart === 'newSection'
          ? sectionCounters.get(region.sectionOccurrenceId) ?? numbering.start
          : sectionCounters.get(region.sectionOccurrenceId)
            ?? continuousCounter
            ?? numbering?.start
            ?? 1;
      counter ??= 1;
      for (const index of indices) {
        let node: PaintNode = body[index]!;
        if (node.kind === 'paragraph' && node.ordinaryFlow && numbering) {
          const numbered = numberedParagraph(
            node,
            counter,
            Math.max(1, numbering.countBy),
            wordLineNumberDistancePt(numbering.distance),
            session,
          );
          node = numbered.paragraph;
          counter = numbered.counterEnd;
        }
        if (translationPt !== 0
          && (node.ordinaryFlow || node.sectionFlowOwnership === 'host-flow')
          && (node.kind === 'paragraph' || node.kind === 'table')) {
          node = translateBodyOccurrence(node, { xPt: 0, yPt: translationPt });
        }
        body[index] = node;
      }
      if (numbering) {
        sectionCounters.set(region.sectionOccurrenceId, counter);
        continuousCounter = counter;
      }
    }
    return Object.freeze({
      ...page,
      layers: replacePageLayerNodes(page.layers, 'body', body),
    });
  });
  return Object.freeze({ ...layout, pages: Object.freeze(pages) });
}
