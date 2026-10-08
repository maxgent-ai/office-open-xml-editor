import { DEFAULT_KINSOKU_RULES } from '@silurus/ooxml-core';
import { layoutLines, type LayoutSeg } from '../line-layout.js';
import type { ParagraphLayoutContext } from '../layout-context.js';
import type { MeasuredParagraph } from '../paragraph-measure.js';
import { paragraphLayoutFromMeasurement } from '../layout/paragraph.js';
import { createTextLayoutService } from '../layout/text.js';
import { createFontResolver } from '../layout/font-service.js';
import { createPageLayers } from '../layout/page-graph.js';
import type { DocumentLayout, ParagraphLayout } from '../layout/types.js';

/** Real breaking and retained placement around an exclusion, with deterministic
 * glyph measurements. Two words occupy separate gaps of each physical line. */
export function gapParagraph(physicalLines = 3, justify = false, pictureInSecondGap = false) {
  const ctx = { measureText: (text: string) => ({ width: text.length * 5,
    fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2 }) } as unknown as CanvasRenderingContext2D;
  const textLayoutService = createTextLayoutService({ fonts: createFontResolver([]),
    measurer: { fingerprint: 'gap-metrics', measure: request => ({
      advancePt: request.text.length * 5, ascentPt: 8, descentPt: 2,
    }) },
  });
  const segments: LayoutSeg[] = Array.from({ length: physicalLines * 2 }, (_, index) => ({
    text: String.fromCharCode(65 + index).repeat(4) + ' ', sourceRunIndex: index,
    textLayoutService, textShapeRequest: { text: String.fromCharCode(65 + index).repeat(4) + ' ',
      fontSizePt: 10, fonts: { ascii: 'Arial' } },
    fontSize: 10, fontFamily: 'Arial', bold: false, italic: false,
    underline: false, strikethrough: false, color: null, vertAlign: null, measuredWidth: 0,
  }));
  if (pictureInSecondGap) segments[1] = {
    imagePath: 'word/media/picture.png', mimeType: 'image/png', widthPt: 20, heightPt: 10,
    anchor: false, anchorXPt: 0, anchorYPt: 0, anchorXFromMargin: false, anchorYFromPara: false,
    inlinePicture: true, sourceRunIndex: 1, measuredWidth: 0,
  };
  const raw = layoutLines(ctx, segments, 200, 0, 1, [], {
    floats: [{ kind: 'shape', mode: 'square', authoredWrap: 'square', imageKey: 'obstacle',
      imageX: 40, imageY: 0, imageW: 120, imageH: 100,
      xLeft: 40, xRight: 160, yTop: 0, yBottom: 100, side: 'bothSides',
      distLeft: 0, distRight: 0, distTop: 0, distBottom: 0, paraId: 0 }],
    paraX: 0, startPageY: 0, columnXPt: 0, columnWidthPt: 200, pageH: 800,
    lineBoxH: () => 10,
  });
  const context: ParagraphLayoutContext = {
    lineGrid: { active: false, pitchPt: null },
    characterGrid: { active: false, kind: null, pitchPt: null, deltaPt: 0 },
    rightIndentGrid: { pitchPt: null, paragraphAllowsAdjustment: true },
    physicalIndentLeftPt: 0, physicalIndentRightPt: 0, firstIndentPt: 0,
    lineSpacing: null, spaceBeforePt: 0, spaceAfterPt: 0,
    baseRtl: false, isJustified: justify, stretchLastLine: false,
    tabStops: [], hasRuby: false, hasEastAsianText: false,
    kinsoku: DEFAULT_KINSOKU_RULES, defaultTabPt: 36,
  };
  const placement = { startYPt: 0, paragraphXPt: 0, availableWidthPt: 200,
    maximumYPt: 800, suppressSpaceBefore: false };
  const measured: MeasuredParagraph = {
    lines: raw.map(layout => ({ layout, topYPt: layout.topY ?? 0, advancePt: 10 })),
    markOnly: false, requestedSpaceBeforePt: 0, requestedSpaceAfterPt: 0,
    uniformRubyAdvancePt: 0, contentStartYPt: 0, contentEndYPt: physicalLines * 10,
    lastLineBelowBaselinePt: 2, placement,
  };
  const paragraph = paragraphLayoutFromMeasurement({
    alignment: justify ? 'both' : 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: segments.map((segment, index) => ({ ...segment, type: 'text' in segment ? 'text' : 'image',
      ...(index === 1 || index === 3 ? { noteRef: { kind: 'footnote', id: String(index) } } : {}),
    })),
  } as never, {
    id: 'gap-paragraph', source: { story: 'body', storyInstance: 'body', path: [0] },
    flowDomainId: 'body', ordinaryFlow: true, context, placement, measurer: {} as never,
    environment: { documentHasEastAsianText: false, pageWritingMode: 'horizontal-tb', pageIndex: 0, totalPages: 1 }, exclusions: [],
  }, measured);
  const boundaries = raw.flatMap((line, index) => raw[index + 1]?.topY === line.topY ? [] : [line.consumedEnd!]);
  return { paragraph, boundaries };
}

export function gapDocument(paragraph: ParagraphLayout): DocumentLayout {
  return { pages: [{
    pageIndex: 0, flowDomains: [], readingOrder: [paragraph.id],
    geometry: { widthPt: 200, heightPt: 800 },
    sectionRegions: [{ id: 'region', sectionOccurrenceId: 'section', flowDomainIds: ['body'],
      blockStartPt: 0, blockEndPt: 800,
      section: { verticalAlignment: 'top', lineNumbering: { start: 1, countBy: 1, restart: 'newPage', distance: 0 } },
    }],
    layers: createPageLayers([{ layer: 'body', node: paragraph, coordinateSpace: 'upright-physical' }]),
  }] } as unknown as DocumentLayout;
}
