import { defineCompatibilityRule } from './compatibility.js';
import type { ParagraphLayout, StoryLayout } from './types.js';

/** Evidence: Word for Mac PDF exports of synthetic controls (issue #1615):
 * margin-, page- and paragraph-relative pictures; square, full-width square
 * and topAndBottom wrap; allowOverlap on and off; single anchors, anchors in
 * adjacent paragraphs and a later anchor sharing a page with an accepted one;
 * multi-line anchor paragraphs with the anchor run on the first or last line;
 * widow control; keepNext on the preceding paragraph; a picture taller than
 * any page remainder; two newspaper columns; a full-width picture mid-body;
 * and a 400-paragraph document with six clustered anchors. */
export const WORD_PAGE_ANCHOR_LINE_DEFERRAL = defineCompatibilityRule({
  id: 'word-page-anchor-line-deferral',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'page-anchor-line-deferral',
    application: 'Microsoft Word',
    version: '16.113.2',
    platform: 'macOS 27.0',
  },
  description: 'ECMA-376 does not say where a page-owned drawing goes when its own wrap exclusion pushes its anchor line off the page. Word tests anchor lines in flow order on each page: with the page-owned anchors already accepted on page N whose lines precede line L, plus the anchors on L, Word lays out N again; if L stays on N (any column of the region) those anchors are accepted and the earlier lines wrap around them. Otherwise page N keeps its layout without them, L starts a later page as if the page ended just above L (keepNext and widow control act as for an overflow; a mid-paragraph anchor line splits its paragraph), and the anchors go to the page where L lands. Counterexamples in the same controls: a line that still fits keeps its picture on N, and the first line of a page stays even when its picture pushes it past the body bottom. Unmeasured: a deferral whose line would otherwise fall in an earlier column of a multi-column page moves to the next page, not the next column.',
});

export const WORD_PAGE_ANCHOR_REGION_REGISTRATION = defineCompatibilityRule({
  id: 'word-page-anchor-region-registration',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'page-anchor-line-deferral',
    application: 'Microsoft Word',
    version: '16.113.2',
    platform: 'macOS 27.0',
  },
  description: 'A page-owned drawing belongs to its page, not to the newspaper column its anchor line reaches: in a two-column control Word keeps column 1 wrapped below a margin-relative topAndBottom picture over column 1 while the anchor paragraph moves into column 2. Register such drawings where their section region opens on the page.',
});

export const WORD_ZERO_RELATIVE_SIZE_EXTENT_FALLBACK = defineCompatibilityRule({
  id: 'word-zero-relative-size',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/layout/anchor-frame.test.ts#uses wp:extent when Word does not support an exact-zero relative size',
  },
  description: 'Word 2010 accepts only positive wp14:pctWidth and wp14:pctHeight values under [MS-ODRAWXML] notes 125/126. Preserve an authored zero as acquisition evidence while resolving the object from wp:extent.',
});

export const WORD_VERTICAL_SECTION_PHYSICAL_DRAWING_LAYER = defineCompatibilityRule({
  id: 'word-vertical-section-physical-drawing-layer',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/anchor-vertical-physical.test.ts#lands an upright-section anchor at the recorded physical centroid',
  },
  description: 'Resolve anchored drawings in an upright vertical section against the physical page frame independently of the rotated text-flow coordinate space.',
});

export const WORD_PAGE_LEVEL_FLOAT_PRESCAN = defineCompatibilityRule({
  id: 'word-page-level-float-prescan',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/page-anchor-prescan.test.ts#pre-scan REGISTERS a page-level (relativeFrom="margin") wrap float on an earlier-scanned paragraph',
  },
  description: 'A wrapping drawing whose vertical reference is page-level participates from page start so source-earlier paragraphs on that page see its exclusion.',
});

export const WORD_PARAGRAPH_ANCHOR_PRE_SPACING_ORIGIN = defineCompatibilityRule({
  id: 'word-paragraph-anchor-pre-spacing-origin',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/anchor-paragraph-spacebefore.test.ts#anchors a wrapSquare paragraph float at the pre-spaceBefore paragraph top',
  },
  description: 'Resolve a paragraph-relative anchored drawing from the paragraph top before applying the paragraph spaceBefore contribution.',
});

export const WORD_VERTICAL_SECTION_PHYSICAL_HEADER_FOOTER = defineCompatibilityRule({
  id: 'word-vertical-section-physical-header-footer',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/vertical-header-footer.test.ts#recovers the physical page box + margins from the logical (swapped) section',
  },
  description: 'Paint a vertical section header and footer in the unrotated physical page frame rather than rotating them with the body text flow.',
});

export const WORD_FRAME_AUTO_WRAP_AROUND = defineCompatibilityRule({
  id: 'word-frame-auto-wrap-around',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/frame-geometry.test.ts#wrap="around" and "auto" → square float (auto ≡ around in Word)',
  },
  description: 'Resolve an authored frame wrap value of auto through the same square side-wrap path as around.',
});

export const WORD_LOWER_LAYER_SAME_PARAGRAPH_ANCHOR_COMPOSITION = defineCompatibilityRule({
  id: 'word-lower-layer-same-paragraph-anchor-composition',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'lower-layer-same-paragraph-anchor-composition',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'Word preserves a source-later, lower-z, page-owned drawing at its authored position when it belongs to the same anchor paragraph as already composed higher layers. This is a Word-observed compatibility override to ECMA-376 §20.4.2.3, not a normative OOXML rule.',
});

export const WORD_TEXTBOX_VISIBLE_ANCHOR_EXTENT = defineCompatibilityRule({
  id: 'word-textbox-visible-anchor-extent',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'textbox-visible-anchor-extent',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'For DrawingML middle and bottom text anchoring, derive the positioned extent through the last visible retained block while preserving structural trailing empty paragraphs and terminal paragraph spacing in the complete story.',
});

export const WORD_OVERLAPPING_LAYOUT_IN_CELL_OVERLAY = defineCompatibilityRule({
  id: 'word-overlapping-layout-in-cell-overlay',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/table-cell-anchor-reflow.test.ts#does not grow an automatic row for an overlapping layoutInCell wrapNone object',
  },
  description: 'Word leaves an overlap-permitted, non-wrapping layoutInCell drawing as an overlay instead of growing its automatic table row. This is an Office compatibility exception to the general resize behavior in ECMA-376 §20.4.2.3 layoutInCell; non-overlap drawings retain normative cell containment.',
});

/** Compatibility projection governed by
 * {@link WORD_OVERLAPPING_LAYOUT_IN_CELL_OVERLAY}. */
export function wordLayoutInCellOwnsRowContainment(
  allowOverlap: boolean,
  wrapKind: 'none' | 'square' | 'tight' | 'through' | 'topAndBottom',
): boolean {
  return !allowOverlap || wrapKind !== 'none';
}

function paragraphContributesTextBoxAnchorExtent(paragraph: ParagraphLayout): boolean {
  if (
    paragraph.shading
    || paragraph.borders.length > 0
    || paragraph.resources.length > 0
    || paragraph.drawings.length > 0
    || paragraph.textBoxes.length > 0
    || paragraph.lineNumbers?.some((line) => line.paintOps.length > 0)
  ) {
    return true;
  }
  return paragraph.lines.some((line) => line.placements.some((placement) => {
    if (placement.kind === 'text' || placement.kind === 'resource' || placement.kind === 'drawing') {
      return true;
    }
    return placement.kind === 'tab' && (placement.leaderGlyphs?.length ?? 0) > 0;
  }));
}

/** Compatibility projection governed by {@link WORD_TEXTBOX_VISIBLE_ANCHOR_EXTENT}. */
export function wordTextBoxVisibleAnchorExtentPt(story: StoryLayout): number {
  const startPt = story.flowBounds.yPt;
  let visibleEndPt: number | undefined;
  for (const block of story.blocks) {
    if (block.kind === 'table') {
      visibleEndPt = Math.max(
        visibleEndPt ?? startPt,
        block.flowBounds.yPt + block.advancePt,
      );
      continue;
    }
    if (block.kind !== 'paragraph' || !paragraphContributesTextBoxAnchorExtent(block)) continue;
    visibleEndPt = Math.max(
      visibleEndPt ?? startPt,
      block.flowBounds.yPt + Math.max(0, block.advancePt - block.spacing.afterPt),
    );
  }
  return visibleEndPt === undefined ? 0 : Math.max(0, visibleEndPt - startPt);
}

export function wordZeroRelativeSizeUsesExtent(fraction: number): boolean {
  return fraction === 0;
}

export function wordPageLevelAnchorY(
  relativeFrom: string | null | undefined,
  paragraphRelativeFallback: boolean,
): boolean {
  if (relativeFrom == null) return !paragraphRelativeFallback;
  return relativeFrom !== 'paragraph'
    && relativeFrom !== 'line'
    && relativeFrom !== 'character';
}

export function wordPreservesLowerLayerSameParagraphComposition(
  movingOwnership: 'page' | 'host',
  movingRelativeHeight: number | null,
  blockerRelativeHeight: number | undefined,
): boolean {
  return movingOwnership === 'page'
    && movingRelativeHeight !== null
    && blockerRelativeHeight !== undefined
    && movingRelativeHeight < blockerRelativeHeight;
}
