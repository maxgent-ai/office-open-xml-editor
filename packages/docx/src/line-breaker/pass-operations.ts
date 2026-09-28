import { LineMeasurementAdapter } from './measurement-adapter.js';
import { graphemeClusterOffsets } from '@silurus/ooxml-core';
import {
  MIN_LINE_GAP,
  prepareFloatWrap,
  computePreparedLineFloatWindow,
  type PreparedFloatWrap,
} from '../float-layout.js';
import { calcEffectiveFontPx, EAST_ASIAN_RE } from '../layout/text.js';
import {
  wordSnapToCharsEastAsianCellCount,
  wordIdeographicSpaceLineEndAllowanceCount,
  wordUniformRunPositionPaintPt,
} from '../layout/line-compatibility.js';
import {
  type LayoutImageSeg,
  type LayoutMathSeg,
  type LayoutSeg,
  type LayoutTabSeg,
  type LayoutTextSeg,
  type LineBoundary,
} from './model.js';
import { createLineBreakerState } from './break-queue.js';
import { applyBidiTabPostPass } from './tabs.js';
import {
  eastAsianGridCountSinglePx,
  measuredLineMetrics,
  nativeCanvasLineRatio,
} from './line-metrics.js';
import {
  RESET_SLICED_TEXT_MEASUREMENT,
  charScaleFactor,
  charSpacingDeltaPx,
  protectedNoBreakOffsets,
  segAdvanceWidth,
  segmentCharacterGridDeltaPx,
  slicedPunctuationCompressions,
  slicedTextMetadata,
  snapToCharsAllocatedWidthPx,
  snapToCharsClass,
} from './advance.js';
import {
  buildFont,
  segmentEastAsiaFloorSingleLinePx,
  segmentIntendedSingleLinePx,
} from './font-routes.js';
import { rubyAscentReservePx } from './ruby-metrics.js';
import { fitCJKPrefix, hasEastAsianVisiblePredecessor } from './fit-search.js';
import { rebaseSeaBreaks } from './text-runs.js';
import {
  keepLeadingKinsoku,
  retractLeadingKinsoku,
  type CrossRunKinsokuRetraction,
} from './kinsoku.js';

import type { LineBreakerPassInput } from './pass-driver.js';

export interface PassOperationState extends LineBreakerPassInput {
  readonly breakerState: ReturnType<typeof createLineBreakerState>;
  readonly sameLatinSpaceFace: (candidate: LayoutTextSeg, reference: LayoutTextSeg) => boolean;
  readonly materializeLatinSpaceCompression: () => void;
  readonly snapPitchPx: number | null;
  readonly minLineStartWidth: () => number;
  readonly isParagraphMarkOnlyFlow: boolean;
  readonly startLine: (minWidth?: number) => void;
  readonly availW: () => number;
  readonly fitsMeasuredWidth: (used: number, available: number) => boolean;
  readonly bidiCustomStopsPx: {
    pos: number;
    alignment: 'left' | 'start' | 'center' | 'right' | 'end' | 'decimal' | 'bar' | 'clear' | 'num';
    leader: 'none' | 'dot' | 'hyphen' | 'underscore' | 'heavy' | 'middleDot';
  }[];
  readonly bidiIntervalPx: number;
  readonly flush: (forceHeight?: number, brTerminated?: boolean, nextStart?: LineBoundary) => void;
  readonly prospectiveSnapAdvance: (s: LayoutTextSeg, naturalWidth: number) => number;
  readonly addToLine: (
    s: LayoutTextSeg | LayoutImageSeg | LayoutMathSeg | LayoutTabSeg,
    w: number,
    h: number,
    asc: number,
    desc: number,
  ) => void;
  readonly measurement: LineMeasurementAdapter;
  readonly effectiveFontPx: (s: LayoutTextSeg) => number;
  readonly measureText: (s: LayoutTextSeg, clusterGeometry?: boolean) => TextMetrics;
  readonly verticalInkExtra: (s: LayoutTextSeg, text: string) => number;
  readonly setMeasureFont: (font: string) => void;
  readonly endBoundary: LineBoundary;
  readonly segNaturalAdvance: (s: LayoutTextSeg) => number;
  readonly standaloneSnapAdvance: (s: LayoutTextSeg, naturalWidth: number) => number;
  readonly segAdvance: (s: LayoutTextSeg) => number;
  readonly strNaturalAdvance: (
    s: LayoutTextSeg,
    text: string,
    retainTrailingPunctuationCompression?: boolean,
  ) => number;
  readonly eastAsianSnapCellCount: (s: LayoutTextSeg) => number;
  readonly strAdvance: (
    s: LayoutTextSeg,
    text: string,
    retainTrailingPunctuationCompression?: boolean,
  ) => number;
  readonly fitHomogeneousLatinSpaces: (next: LayoutTextSeg, nextFitWidth: number) => boolean;
  readonly textSegmentBox: (
    s: LayoutTextSeg,
  ) => Readonly<{ width: number; height: number; ascent: number; descent: number }>;
  readonly appendQueuedIdeographicSpaceSegment: (source: LayoutTextSeg) => void;
  readonly tabFollowWidth: (q: LayoutSeg) => number;
  readonly decimalAlignmentPoint: (
    segments: readonly LayoutSeg[],
  ) => Readonly<{ segmentIndex: number; charOffset: number }> | null;
  readonly decimalAlignmentPrefixWidth: (segments: readonly LayoutSeg[]) => number | undefined;
  readonly tabFollowingMetrics: () => Readonly<{ totalWidth: number; decimalPrefixWidth?: number }>;
  readonly emergencyTextSplit: (
    segment: LayoutTextSeg,
    available: number,
    forceAtLeastOne?: boolean,
  ) => number;
  readonly externalLinkSyntaxSplit: (segment: LayoutTextSeg, available: number) => number;
  readonly queueEmergencyTail: (segment: LayoutTextSeg, split: number) => void;
  readonly retractCurrentLineForLeadingKinsoku: (next: LayoutTextSeg) => CrossRunKinsokuRetraction;
  readonly keepLeadingKinsokuWithCurrentLine: (
    segment: LayoutTextSeg,
    h: number,
    asc: number,
    desc: number,
  ) => boolean;
  readonly probeHeights: readonly number[] | null;
  readonly preparedFloatWrap?: PreparedFloatWrap;
}

export function performSameLatinSpaceFace(
  candidate: LayoutTextSeg,
  reference: LayoutTextSeg,
): boolean {
  return (
    candidate.latinSpaceCompressionEligible === true &&
    !candidate.verticalRun &&
    !reference.verticalRun &&
    !candidate.tateChuYoko &&
    !reference.tateChuYoko &&
    candidate.latinSpaceAverageWidthRatio === reference.latinSpaceAverageWidthRatio &&
    candidate.fontRoute?.fingerprint === reference.fontRoute?.fingerprint &&
    candidate.fontFamily === reference.fontFamily &&
    candidate.fontSize === reference.fontSize &&
    candidate.bold === reference.bold &&
    candidate.italic === reference.italic &&
    (candidate.charScale == null || candidate.charScale === 1) &&
    (reference.charScale == null || reference.charScale === 1) &&
    candidate.kerning === reference.kerning &&
    candidate.widthBalanceGridDeltaFactor === reference.widthBalanceGridDeltaFactor &&
    !candidate.rtl &&
    candidate.fitTextRegionIndex === undefined
  );
}

export function performMaterializeLatinSpaceCompression(operationState: PassOperationState): void {
  const { breakerState } = operationState;

  for (let index = 0; index < breakerState.latinAppliedGapCount; index += 1) {
    const gap = breakerState.latinLineGaps[index];
    gap.measuredWidth -= breakerState.latinAppliedPerGap;
    gap.latinSpaceCompressionPx = breakerState.latinAppliedPerGap;
  }
  breakerState.latinAppliedGapCount = 0;
  breakerState.latinAppliedPerGap = 0;
}

export function performStartLine(operationState: PassOperationState, minWidth: number = 0): void {
  const { breakerState, maxWidth, wrapCtx, baseRtl, probeHeights, preparedFloatWrap } =
    operationState;

  breakerState.snapBlock = null;
  breakerState.lineXOffset = 0;
  breakerState.lineMaxWidth = maxWidth;
  if (!wrapCtx) return;
  const probeH = probeHeights?.[breakerState.lines.length];
  // The first pass measures this line without a float window. A later pass
  // resolves it only once that exact line index has an observed line-box
  // height; newly-created lines are likewise measured before they are probed.
  if (probeH === undefined) return;
  const reference = {
    xLeftPt: wrapCtx.referenceXPt ?? wrapCtx.paraX,
    xRightPt: (wrapCtx.referenceXPt ?? wrapCtx.paraX) + (wrapCtx.referenceWidthPt ?? maxWidth),
    readingDirection: wrapCtx.readingDirection ?? (baseRtl ? 'rtl' : 'ltr'),
  } as const;
  if (wrapCtx.lineWindow) {
    const win = wrapCtx.lineWindow({
      topYPt: breakerState.currentLineTopY,
      minimumStartWidthPt: MIN_LINE_GAP,
      squareMinimumStartWidthPt: minWidth,
      probeHeightPt: probeH,
      paragraphXPt: wrapCtx.paraX,
      maximumWidthPt: maxWidth,
      columnXPt: wrapCtx.columnXPt,
      columnWidthPt: wrapCtx.columnWidthPt,
    });
    breakerState.currentLineTopY = win.topYPt;
    breakerState.lineXOffset = win.xOffsetPt;
    breakerState.lineMaxWidth = win.maximumWidthPt;
  } else {
    const win = computePreparedLineFloatWindow(
      breakerState.currentLineTopY,
      MIN_LINE_GAP,
      probeH,
      wrapCtx.paraX,
      maxWidth,
      preparedFloatWrap ?? prepareFloatWrap(wrapCtx.floats),
      wrapCtx.columnXPt,
      wrapCtx.columnXPt + wrapCtx.columnWidthPt,
      reference,
      minWidth,
    );
    breakerState.currentLineTopY = win.topY;
    breakerState.lineXOffset = win.xOffset;
    breakerState.lineMaxWidth = win.maxWidth;
  }
}

export function performAvailW(operationState: PassOperationState) {
  const { breakerState, firstIndent, widthPolicy } = operationState;
  return widthPolicy === 'intrinsic'
    ? Number.POSITIVE_INFINITY
    : breakerState.lineMaxWidth - (breakerState.isFirst ? firstIndent : 0);
}

export function performFitsMeasuredWidth(
  operationState: PassOperationState,
  used: number,
  available: number,
): boolean {
  const { breakerState, firstIndent, widthPolicy } = operationState;

  if (used <= available) return true;
  if (
    widthPolicy === 'intrinsic' ||
    !Number.isFinite(used) ||
    !Number.isFinite(breakerState.lineMaxWidth)
  ) {
    return false;
  }
  return used + (breakerState.isFirst ? firstIndent : 0) <= breakerState.lineMaxWidth;
}

export function performFlush(
  operationState: PassOperationState,
  forceHeight?: number,
  brTerminated = false,
  nextStart?: LineBoundary,
) {
  const {
    breakerState,
    materializeLatinSpaceCompression,
    minLineStartWidth,
    startLine,
    bidiCustomStopsPx,
    bidiIntervalPx,
    endBoundary,
    strAdvance,
    decimalAlignmentPoint,
    firstIndent,
    scale,
    wrapCtx,
    tabOriginPx,
    marginRightPx,
    baseRtl,
  } = operationState;

  materializeLatinSpaceCompression();
  breakerState.currentWidth += applyBidiTabPostPass({
    baseRtl,
    currentLine: breakerState.currentLine,
    marginRightPx,
    lineXOffset: breakerState.lineXOffset,
    lineMaxWidth: breakerState.lineMaxWidth,
    isFirst: breakerState.isFirst,
    firstIndent,
    tabOriginPx,
    bidiCustomStopsPx,
    bidiIntervalPx,
    decimalAlignmentPoint,
    strAdvance,
  });
  // §17.3.2.24 defines `position` relative to surrounding non-positioned
  // text. A line whose every metric-bearing item shares the same inherited
  // position has no differently-positioned peer to pin the resulting line
  // box to one side. `word-uniform-run-position-leading` owns the compatibility
  // placement of that box around the glyphs. Keep mixed
  // lines relative to zero so their authored displacement and ink union
  // remain unchanged. Images/math
  // provide a zero-position reference; tabs do not contribute vertical
  // metrics. The fixed drop-cap path intentionally keeps its paint-only
  // lowering and therefore opts out of this normalization.
  let commonPositionPt: number | undefined;
  let hasPositionReference = false;
  for (const segment of breakerState.currentLine) {
    if ('isTab' in segment) continue;
    const positionPt = 'text' in segment ? (segment.position ?? 0) : 0;
    if ('text' in segment && segment.positionExtendsLineBox === false) {
      commonPositionPt = 0;
      hasPositionReference = true;
      break;
    }
    if (!hasPositionReference) {
      commonPositionPt = positionPt;
      hasPositionReference = true;
    } else if (commonPositionPt !== positionPt) {
      commonPositionPt = 0;
      break;
    }
  }
  const linePositionReferencePt = hasPositionReference ? (commonPositionPt ?? 0) : 0;
  if (linePositionReferencePt !== 0) {
    for (const segment of breakerState.currentLine) {
      if ('text' in segment) {
        segment.lineRelativePosition = wordUniformRunPositionPaintPt(
          segment.position ?? 0,
          linePositionReferencePt,
        );
      }
    }
  }
  // §17.3.3.1 — the break is one run among the line's runs: its own size
  // participates in the line height but must not override a taller peer.
  const h =
    forceHeight !== undefined
      ? Math.max(breakerState.lineHeight, forceHeight)
      : breakerState.lineHeight || 10;
  // If the line has no measured content (empty/line-break line), synthesize
  // stable ascent/descent from the effective font size so wrap/baseline math
  // stays consistent with non-empty lines.
  const hasContent = breakerState.lineAscent > 0 || breakerState.lineDescent > 0;
  const asc = hasContent ? breakerState.lineAscent : h * scale * 0.8;
  const desc = hasContent ? breakerState.lineDescent : h * scale * 0.2;
  const visibleAscent = breakerState.lineHasVisibleMetrics ? breakerState.lineVisibleAscent : asc;
  const visibleDescent = breakerState.lineHasVisibleMetrics
    ? breakerState.lineVisibleDescent
    : desc;
  const visibleIntendedSingle = breakerState.lineHasVisibleMetrics
    ? breakerState.lineVisibleIntendedSingle
    : breakerState.lineIntendedSingle;
  const gridCountSingle =
    breakerState.lineGridCountSingle ||
    (breakerState.lineEastAsian
      ? eastAsianGridCountSinglePx(breakerState.lineIntendedSingle, h * scale)
      : asc + desc);
  const inlinePictureTextSingle = breakerState.lineHasInlinePicture
    ? Math.max(breakerState.lineIntendedSingle, breakerState.linePictureMarkSingle)
    : 0;
  // Only project that registered rule when every metric-bearing item is
  // visible text in one admitted face tuple. Canvas fallback geometry does
  // not reveal hhea descent, and mixed styles cannot share one descent
  // reserve. The rule uses face data, never a family-specific correction.
  const positionedTexts = breakerState.currentLine.filter(
    (segment): segment is LayoutTextSeg => 'text' in segment,
  );
  const firstPositioned = positionedTexts[0];
  const uniformPositionAuto =
    linePositionReferencePt !== 0 &&
    positionedTexts.length > 0 &&
    breakerState.currentLine.every((segment) => 'isTab' in segment || 'text' in segment) &&
    firstPositioned?.resolvedDesignDescentRatio != null &&
    (firstPositioned.referenceFontVerticalMetric ||
      firstPositioned.resolvedResourceVerticalMetric) &&
    positionedTexts.every(
      (segment) =>
        segment.text.length > 0 &&
        !segment.metricOnly &&
        !segment.ruby &&
        !segment.vertAlign &&
        segment.positionExtendsLineBox !== false &&
        segment.position === linePositionReferencePt &&
        segment.fontFamily === firstPositioned.fontFamily &&
        segment.fontRoute?.fingerprint === firstPositioned.fontRoute?.fingerprint &&
        segment.bold === firstPositioned.bold &&
        segment.italic === firstPositioned.italic &&
        segment.fontSize === firstPositioned.fontSize &&
        segment.resolvedDesignDescentRatio === firstPositioned.resolvedDesignDescentRatio &&
        segment.referenceFontVerticalMetric === firstPositioned.referenceFontVerticalMetric &&
        segment.resolvedResourceVerticalMetric === firstPositioned.resolvedResourceVerticalMetric &&
        (segment.referenceFontVerticalMetric || segment.resolvedResourceVerticalMetric),
    )
      ? {
          normalSinglePx: Math.max(
            asc + desc - Math.abs(linePositionReferencePt * scale),
            breakerState.lineIntendedSingle,
          ),
          positionPx: linePositionReferencePt * scale,
          designDescentPx:
            firstPositioned.resolvedDesignDescentRatio * firstPositioned.fontSize * scale,
        }
      : undefined;
  breakerState.lines.push({
    segments: breakerState.currentLine,
    height: h,
    ascent: asc,
    descent: desc,
    visibleAscent,
    visibleDescent,
    visibleIntendedSingle,
    intendedSingle: breakerState.lineIntendedSingle,
    ...(inlinePictureTextSingle > 0 ? { inlinePictureTextSingle } : {}),
    uniformPositionAuto,
    // Empty/synthetic East Asian lines use the same design-height rule as a
    // text run; their synthesized Canvas box must not reintroduce a
    // scale-dependent cell count.
    gridCountSingle,
    xOffset: breakerState.lineXOffset,
    availWidth: breakerState.lineMaxWidth,
    topY: wrapCtx ? breakerState.currentLineTopY : undefined,
    hasRuby: breakerState.lineHasRuby,
    eastAsian: breakerState.lineEastAsian,
    endsWithBreak: brTerminated,
    consumedEnd: nextStart ?? breakerState.queue[0]?.src ?? endBoundary,
  });
  if (wrapCtx) {
    breakerState.currentLineTopY += wrapCtx.lineBoxH(
      asc,
      desc,
      breakerState.lineHasRuby,
      breakerState.lineIntendedSingle,
      breakerState.lineEastAsian,
      gridCountSingle,
      uniformPositionAuto,
      inlinePictureTextSingle,
    );
  }
  breakerState.currentLine = [];
  breakerState.currentWidth = 0;
  breakerState.latinLineFace = undefined;
  breakerState.latinLineHomogeneous = true;
  breakerState.latinLineGaps = [];
  breakerState.latinUniformGapCapacity = undefined;
  breakerState.lineHeight = 0;
  breakerState.lineAscent = 0;
  breakerState.lineDescent = 0;
  breakerState.lineIntendedSingle = 0;
  breakerState.lineHasInlinePicture = false;
  breakerState.linePictureMarkSingle = 0;
  breakerState.lineGridCountSingle = 0;
  breakerState.lineVisibleAscent = 0;
  breakerState.lineVisibleDescent = 0;
  breakerState.lineVisibleIntendedSingle = 0;
  breakerState.lineHasVisibleMetrics = false;
  breakerState.lineHasRuby = false;
  breakerState.lineEastAsian = false;
  breakerState.isFirst = false;
  startLine(minLineStartWidth());
}

export function performProspectiveSnapAdvance(
  operationState: PassOperationState,
  s: LayoutTextSeg,
  naturalWidth: number,
): number {
  const { breakerState, snapPitchPx, eastAsianSnapCellCount, characterGrid } = operationState;

  const kind = snapToCharsClass(s, characterGrid);
  if (!kind || snapPitchPx == null) return naturalWidth;
  if (kind === 'eastAsia') {
    const cells = eastAsianSnapCellCount(s);
    return snapToCharsAllocatedWidthPx(naturalWidth, kind, snapPitchPx, cells);
  }
  if (breakerState.snapBlock?.kind === kind) {
    return (
      snapToCharsAllocatedWidthPx(
        breakerState.snapBlock.naturalWidthPx + naturalWidth,
        kind,
        snapPitchPx,
      ) - breakerState.snapBlock.allocatedWidthPx
    );
  }
  return snapToCharsAllocatedWidthPx(naturalWidth, kind, snapPitchPx);
}

export function performAddToLine(
  operationState: PassOperationState,
  s: LayoutTextSeg | LayoutImageSeg | LayoutMathSeg | LayoutTabSeg,
  w: number,
  h: number,
  asc: number,
  desc: number,
) {
  const {
    breakerState,
    sameLatinSpaceFace,
    materializeLatinSpaceCompression,
    snapPitchPx,
    effectiveFontPx,
    eastAsianSnapCellCount,
    ctx,
    scale,
    fontFamilyClasses,
    characterGrid,
  } = operationState;

  let committedWidth = w;
  if ('text' in s) {
    const kind = snapToCharsClass(s, characterGrid);
    const naturalWidth = s.snapGridNaturalWidthPx ?? w;
    if (kind && snapPitchPx != null) {
      s.snapGridClass = kind;
      s.snapGridNaturalWidthPx = naturalWidth;
      s.snapGridCellPitchPx = snapPitchPx;
      if (kind === 'eastAsia') {
        const cellCount = eastAsianSnapCellCount(s);
        committedWidth = snapToCharsAllocatedWidthPx(naturalWidth, kind, snapPitchPx, cellCount);
        s.snapGridLeadingPadPx = 0;
        s.snapGridTrailingPadPx = committedWidth - naturalWidth;
        s.measuredWidth = committedWidth;
        breakerState.snapBlock = null;
      } else if (breakerState.snapBlock?.kind === kind) {
        const previousLeading = breakerState.snapBlock.first.snapGridLeadingPadPx ?? 0;
        const previousTrailing = breakerState.snapBlock.last.snapGridTrailingPadPx ?? 0;
        const combinedNatural = breakerState.snapBlock.naturalWidthPx + naturalWidth;
        const combinedAllocated = snapToCharsAllocatedWidthPx(combinedNatural, kind, snapPitchPx);
        const slack = combinedAllocated - combinedNatural;
        const leading = kind === 'latin' ? slack / 2 : 0;
        const trailing = slack - leading;
        breakerState.snapBlock.first.measuredWidth -= previousLeading;
        breakerState.snapBlock.first.snapGridLeadingPadPx = leading;
        breakerState.snapBlock.first.measuredWidth += leading;
        breakerState.snapBlock.last.measuredWidth -= previousTrailing;
        s.snapGridLeadingPadPx = 0;
        s.snapGridTrailingPadPx = trailing;
        s.measuredWidth = naturalWidth + trailing;
        committedWidth = combinedAllocated - breakerState.snapBlock.allocatedWidthPx;
        breakerState.snapBlock = {
          kind,
          first: breakerState.snapBlock.first,
          last: s,
          naturalWidthPx: combinedNatural,
          allocatedWidthPx: combinedAllocated,
        };
      } else {
        const allocated = snapToCharsAllocatedWidthPx(naturalWidth, kind, snapPitchPx);
        const slack = allocated - naturalWidth;
        const leading = kind === 'latin' ? slack / 2 : 0;
        const trailing = slack - leading;
        s.snapGridLeadingPadPx = leading;
        s.snapGridTrailingPadPx = trailing;
        s.measuredWidth = allocated;
        committedWidth = allocated;
        breakerState.snapBlock = {
          kind,
          first: s,
          last: s,
          naturalWidthPx: naturalWidth,
          allocatedWidthPx: allocated,
        };
      }
    } else {
      s.snapGridClass = undefined;
      s.snapGridLeadingPadPx = undefined;
      s.snapGridTrailingPadPx = undefined;
      s.snapGridCellPitchPx = undefined;
      s.measuredWidth = w;
      breakerState.snapBlock = null;
    }
  } else {
    breakerState.snapBlock = null;
  }
  breakerState.currentLine.push(s);
  breakerState.currentWidth += committedWidth;
  if (
    'text' in s &&
    s.latinSpaceCompressionEligible === true &&
    s.latinSpaceAverageWidthRatio != null &&
    s.fontRoute
  ) {
    if (breakerState.latinLineFace && !sameLatinSpaceFace(s, breakerState.latinLineFace)) {
      materializeLatinSpaceCompression();
      breakerState.latinLineHomogeneous = false;
    }
    breakerState.latinLineFace ??= s;
    if (s.latinNaturalTrailingSpacePx !== undefined) {
      const floor =
        ((calcEffectiveFontPx(s, scale) * s.latinSpaceAverageWidthRatio) / 2) * charScaleFactor(s) +
        segmentCharacterGridDeltaPx(s, characterGrid, scale);
      const capacity = Math.max(0, s.latinNaturalTrailingSpacePx - floor);
      if (
        breakerState.latinUniformGapCapacity !== undefined &&
        Math.abs(capacity - breakerState.latinUniformGapCapacity) > 1e-6
      ) {
        materializeLatinSpaceCompression();
        breakerState.latinLineHomogeneous = false;
      }
      breakerState.latinUniformGapCapacity ??= capacity;
      breakerState.latinLineGaps.push(s);
    }
  } else {
    materializeLatinSpaceCompression();
    breakerState.latinLineHomogeneous = false;
  }
  if (h > breakerState.lineHeight) breakerState.lineHeight = h;
  if ('imagePath' in s && s.inlinePicture === true) {
    breakerState.lineHasInlinePicture = true;
    breakerState.linePictureMarkSingle = Math.max(
      breakerState.linePictureMarkSingle,
      (s.paragraphMarkSinglePx ?? 0) * scale,
    );
  }
  if (asc > breakerState.lineAscent) breakerState.lineAscent = asc;
  if (desc > breakerState.lineDescent) breakerState.lineDescent = desc;
  const paintsInlineInk = !('text' in s) || s.metricOnly !== true;
  if (paintsInlineInk) {
    breakerState.lineHasVisibleMetrics = true;
    if (asc > breakerState.lineVisibleAscent) breakerState.lineVisibleAscent = asc;
    if (desc > breakerState.lineVisibleDescent) breakerState.lineVisibleDescent = desc;
  }
  // Grid-count height for docGrid cell allocation (§17.6.5). Only East Asian
  // TEXT (and tall inline objects) drives the count — a Latin run keeps its
  // natural height and is NOT cell-rounded, so it must not contribute (its
  // substituted Canvas box would otherwise inflate the count). An EA text run
  // counts from its DESIGN height when tabled, else the deterministic Word FE
  // 1.3em fallback; an image/math object counts its measured box. The line's
  // value is the max.
  let segGridCount = 0;
  if (!('isTab' in s) && !('imagePath' in s) && !('math' in s)) {
    const ts = s as LayoutTextSeg;
    if (ts.ruby) breakerState.lineHasRuby = true;
    const metricEastAsian = ts.metricEastAsian === true || EAST_ASIAN_RE.test(ts.text);
    if (!breakerState.lineEastAsian && metricEastAsian) breakerState.lineEastAsian = true;
    // Prefer the selected resource's single-line height. Without admitted
    // geometry, the generic East Asian grid fallback remains authoritative.
    // Small caps (non-super/sub) keep the FULL run size here so the line box
    // follows the run size, not the 2pt-reduced glyphs (§17.3.2.33).
    const intendedEm = ts.smallCaps && !ts.vertAlign ? ts.fontSize * scale : effectiveFontPx(ts);
    // The OpenType code-page class selects the general line ratio even for
    // Latin text. This script hint selects an optional East-Asian-specific
    // floor and grid-cell counting; ruby keeps its measured annotation box.
    const segScriptHint = metricEastAsian && !ts.ruby;
    const nativeRatio =
      ts.resolvedLineHeightRatio == null
        ? nativeCanvasLineRatio(
            ctx,
            fontFamilyClasses,
            ts.fontRoute,
            ts.fontFamily,
            ts.bold ? 700 : 400,
            ts.italic ? 'italic' : 'normal',
            ts.text,
          )
        : null;
    const designIntended =
      ts.textBoxLineFloor && ts.ruby
        ? 0
        : Math.max(
            segmentIntendedSingleLinePx(ts, intendedEm, segScriptHint),
            ts.textBoxLineFloor || ts.metricEastAsian === true
              ? segmentEastAsiaFloorSingleLinePx(ts, intendedEm, segScriptHint)
              : 0,
          );
    const intended = Math.max(designIntended, (nativeRatio ?? 0) * intendedEm);
    if (intended > breakerState.lineIntendedSingle) breakerState.lineIntendedSingle = intended;
    if (paintsInlineInk && intended > breakerState.lineVisibleIntendedSingle) {
      breakerState.lineVisibleIntendedSingle = intended;
    }
    // Only East Asian text is cell-rounded. The native Canvas probe can
    // establish a browser-selected font box for ordinary auto lines, but it
    // cannot establish Word's Far-East design height or OS/2 code-page class.
    // For an untabled tuple retain the documented 1.3em grid fallback;
    // parsed resource/reference design metrics still take precedence.
    if (segScriptHint) segGridCount = eastAsianGridCountSinglePx(designIntended, intendedEm);
  } else if (!('isTab' in s)) {
    // Image/math object: a tall inline object sizes the line's cells too.
    segGridCount = asc + desc;
  }
  if (segGridCount > breakerState.lineGridCountSingle)
    breakerState.lineGridCountSingle = segGridCount;
}

export function performSegNaturalAdvance(
  operationState: PassOperationState,
  s: LayoutTextSeg,
): number {
  const { measureText, verticalInkExtra, scale, characterGrid } = operationState;
  return segAdvanceWidth(
    s,
    measureText(s).width + verticalInkExtra(s, s.text),
    characterGrid,
    scale,
  );
}

export function performStandaloneSnapAdvance(
  operationState: PassOperationState,
  s: LayoutTextSeg,
  naturalWidth: number,
): number {
  const { snapPitchPx, eastAsianSnapCellCount, characterGrid } = operationState;

  const kind = snapToCharsClass(s, characterGrid);
  if (!kind || snapPitchPx == null || s.text.length === 0) return naturalWidth;
  return snapToCharsAllocatedWidthPx(
    naturalWidth,
    kind,
    snapPitchPx,
    kind === 'eastAsia' ? eastAsianSnapCellCount(s) : 1,
  );
}

export function performSegAdvance(operationState: PassOperationState, s: LayoutTextSeg): number {
  const { segNaturalAdvance, standaloneSnapAdvance } = operationState;
  return standaloneSnapAdvance(s, segNaturalAdvance(s));
}

export function performStrNaturalAdvance(
  operationState: PassOperationState,
  s: LayoutTextSeg,
  text: string,
  retainTrailingPunctuationCompression = false,
): number {
  const { measurement, effectiveFontPx, verticalInkExtra, scale, characterGrid } = operationState;

  const start = retainTrailingPunctuationCompression ? s.text.length - text.length : 0;
  const measuredSegment = {
    ...s,
    text,
    punctuationCompressions: slicedPunctuationCompressions(
      s,
      Math.max(0, start),
      Math.max(0, start) + text.length,
    ),
  };
  if (s.textLayoutService && s.textShapeRequest) {
    const shaped = s.textLayoutService.shape({
      ...s.textShapeRequest,
      text,
      fontSizePt: effectiveFontPx(s),
      measure: true,
      clusterGeometry: false,
    });
    return segAdvanceWidth(
      measuredSegment,
      shaped.advancePt + verticalInkExtra(s, text),
      characterGrid,
      scale,
    );
  }
  const natural = measurement.measureRunText(s, text).width;
  return segAdvanceWidth(
    measuredSegment,
    natural + verticalInkExtra(s, text),
    characterGrid,
    scale,
  );
}

export function performEastAsianSnapCellCount(
  operationState: PassOperationState,
  s: LayoutTextSeg,
): number {
  const { snapPitchPx, measurement, measureText, verticalInkExtra, scale, characterGrid } =
    operationState;

  if (snapPitchPx == null) return 1;
  if (s.textLayoutService && s.textShapeRequest && !s.shapedClusters) {
    measureText(s, true);
  }
  const shapedClusters = s.shapedClusters?.length ? s.shapedClusters : null;
  const boundaries =
    shapedClusters == null
      ? [...new Set([0, ...graphemeClusterOffsets(s.text), s.text.length])].sort((a, b) => a - b)
      : null;
  const ranges =
    shapedClusters?.map((cluster) => ({
      start: cluster.range.start,
      end: cluster.range.end,
      advancePx: cluster.advancePt,
    })) ??
    boundaries!.slice(0, -1).map((start, index) => ({
      start,
      end: boundaries![index + 1]!,
      advancePx: undefined,
    }));
  let cells = 0;
  for (const range of ranges) {
    const { start, end } = range;
    if (end <= start) continue;
    const text = s.text.slice(start, end);
    const measuredSegment = {
      ...s,
      text,
      punctuationCompressions: slicedPunctuationCompressions(s, start, end),
    };
    let naturalAdvancePx: number;
    if (range.advancePx != null) {
      naturalAdvancePx = segAdvanceWidth(
        measuredSegment,
        range.advancePx + verticalInkExtra(s, text),
        characterGrid,
        scale,
      );
    } else {
      const naturalWidthPx = measurement.measureRunText(s, text).width;
      naturalAdvancePx = segAdvanceWidth(
        measuredSegment,
        naturalWidthPx + verticalInkExtra(s, text),
        characterGrid,
        scale,
      );
    }
    cells += wordSnapToCharsEastAsianCellCount(naturalAdvancePx, snapPitchPx);
  }
  return Math.max(1, cells);
}

export function performStrAdvance(
  operationState: PassOperationState,
  s: LayoutTextSeg,
  text: string,
  retainTrailingPunctuationCompression = false,
): number {
  const { standaloneSnapAdvance, strNaturalAdvance } = operationState;

  const candidate = {
    ...s,
    text,
    shapedClusters: text === s.text ? s.shapedClusters : undefined,
  };
  return standaloneSnapAdvance(
    candidate,
    strNaturalAdvance(s, text, retainTrailingPunctuationCompression),
  );
}

export function performFitHomogeneousLatinSpaces(
  operationState: PassOperationState,
  next: LayoutTextSeg,
  nextFitWidth: number,
): boolean {
  const {
    breakerState,
    sameLatinSpaceFace,
    availW,
    fitsMeasuredWidth,
    characterGrid,
    baseRtl,
    isJustified,
    widthPolicy,
  } = operationState;

  if (
    isJustified ||
    baseRtl ||
    widthPolicy !== 'bounded' ||
    characterGrid?.type === 'snapToChars' ||
    (characterGrid?.type === 'linesAndChars' && next.widthBalanceGridDeltaFactor !== 0.5) ||
    next.latinSpaceCompressionEligible !== true ||
    next.latinSpaceAverageWidthRatio == null ||
    !next.fontRoute ||
    next.rtl ||
    next.verticalRun ||
    next.tateChuYoko ||
    next.fitTextRegionIndex !== undefined ||
    !breakerState.latinLineHomogeneous ||
    !breakerState.latinLineFace ||
    !sameLatinSpaceFace(next, breakerState.latinLineFace) ||
    breakerState.latinLineGaps.length === 0
  )
    return false;
  const totalCapacity =
    (breakerState.latinUniformGapCapacity ?? 0) * breakerState.latinLineGaps.length;
  if (totalCapacity <= 0) return false;
  const restored = breakerState.latinAppliedPerGap * breakerState.latinAppliedGapCount;
  const required = Math.max(0, breakerState.currentWidth + restored + nextFitWidth - availW());
  if (
    required > totalCapacity ||
    !fitsMeasuredWidth(breakerState.currentWidth + restored + nextFitWidth - required, availW())
  ) {
    return false;
  }
  // Keep aggregate fit width current. Write each retained gap exactly once
  // when the line is finalized, avoiding quadratic work on long lines.
  breakerState.currentWidth += restored - required;
  breakerState.latinAppliedGapCount = breakerState.latinLineGaps.length;
  breakerState.latinAppliedPerGap = required / breakerState.latinAppliedGapCount;
  return true;
}

export function performTextSegmentBox(
  operationState: PassOperationState,
  s: LayoutTextSeg,
): Readonly<{
  width: number;
  height: number;
  ascent: number;
  descent: number;
}> {
  const {
    measurement,
    effectiveFontPx,
    measureText,
    verticalInkExtra,
    ctx,
    scale,
    fontFamilyClasses,
    characterGrid,
  } = operationState;

  const measured = measureText(s, snapToCharsClass(s, characterGrid) === 'eastAsia');
  const width = segAdvanceWidth(
    s,
    measured.width + verticalInkExtra(s, s.text),
    characterGrid,
    scale,
  );
  s.snapGridNaturalWidthPx = width;

  const fullPx = s.fontSize * scale;
  let metricMeasurement = measured;
  let metricEmPx = effectiveFontPx(s);
  if (s.smallCaps && !s.vertAlign && metricEmPx !== fullPx) {
    if (s.textLayoutService && s.textShapeRequest) {
      const shaped = s.textLayoutService.shape({
        ...s.textShapeRequest,
        text: s.text || 'X',
        fontSizePt: fullPx,
        measure: true,
        clusterGeometry: false,
      });
      metricMeasurement = {
        width: shaped.advancePt,
        actualBoundingBoxAscent: shaped.ascentPt,
        actualBoundingBoxDescent: shaped.descentPt,
        fontBoundingBoxAscent: shaped.ascentPt,
        fontBoundingBoxDescent: shaped.descentPt,
      } as TextMetrics;
    } else {
      metricMeasurement = measurement.measureWithFont(
        buildFont(s.bold, s.italic, fullPx, s.fontFamily, fontFamilyClasses, s.fontRoute),
        s.text || 'X',
      );
    }
    metricEmPx = fullPx;
  }

  const corrected = measuredLineMetrics(metricMeasurement, fullPx);
  // Selected resource sides come from the same admitted face as the line
  // height; native reference sides are policy geometry only. Ruby and an
  // authored baseline position compose additional boxes outside either
  // simple OpenType projection, so retain measured sides for those inputs.
  const designOwnsSides =
    (s.resolvedResourceVerticalMetric || s.referenceFontVerticalMetric) &&
    !s.ruby &&
    (s.position ?? 0) === 0 &&
    s.resolvedDesignAscentRatio != null &&
    s.resolvedDesignDescentRatio != null;
  let ascent = designOwnsSides ? s.resolvedDesignAscentRatio! * metricEmPx : corrected.ascent;
  let descent = designOwnsSides ? s.resolvedDesignDescentRatio! * metricEmPx : corrected.descent;
  if (s.positionExtendsLineBox !== false) {
    const positionPx = (s.position ?? 0) * scale;
    if (positionPx > 0) ascent += positionPx;
    else if (positionPx < 0) descent -= positionPx;
  }
  if (s.ruby && (!s.textBoxLineFloor || s.textBoxVertical)) {
    ascent += rubyAscentReservePx(
      s.ruby.fontSizePt,
      s.ruby.hpsRaisePt,
      scale,
      s,
      ctx,
      fontFamilyClasses,
    );
  }
  return { width, height: s.fontSize, ascent, descent };
}

export function performAppendQueuedIdeographicSpaceSegment(
  operationState: PassOperationState,
  source: LayoutTextSeg,
): void {
  const { breakerState, addToLine, textSegmentBox } = operationState;

  if (
    /\s$/u.test(source.text) ||
    source.ruby !== undefined ||
    source.tateChuYoko === true ||
    source.fitTextRegionIndex !== undefined
  )
    return;
  const follower = breakerState.queue[0];
  if (
    !follower ||
    !('text' in follower) ||
    follower.joinPrev !== true ||
    follower.text.length === 0 ||
    [...follower.text].some((character) => character !== '\u3000')
  )
    return;
  breakerState.queue.shift();
  const hangingCount = wordIdeographicSpaceLineEndAllowanceCount(
    hasEastAsianVisiblePredecessor(source.text),
    follower.paragraphFinalIdeographicSpaceCount ?? [...follower.text].length,
  );
  if (hangingCount === 0) {
    breakerState.queue.unshift(follower);
    return;
  }
  const hangingText = follower.text.slice(0, hangingCount);
  const hangingSegment: LayoutTextSeg = {
    ...follower,
    ...RESET_SLICED_TEXT_MEASUREMENT,
    text: hangingText,
    measuredWidth: 0,
    ...slicedTextMetadata(follower, 0, hangingText.length),
  };
  const followerBox = textSegmentBox(hangingSegment);
  hangingSegment.measuredWidth = followerBox.width;
  addToLine(
    hangingSegment,
    followerBox.width,
    followerBox.height,
    followerBox.ascent,
    followerBox.descent,
  );
  const remainder = follower.text.slice(hangingText.length);
  if (remainder.length > 0) {
    breakerState.queue.unshift({
      ...follower,
      ...RESET_SLICED_TEXT_MEASUREMENT,
      text: remainder,
      measuredWidth: 0,
      joinPrev: undefined,
      hardJoinPrev: undefined,
      ...slicedTextMetadata(follower, hangingText.length, follower.text.length),
      src: follower.src
        ? {
            segIndex: follower.src.segIndex,
            charOffset: follower.src.charOffset + hangingText.length,
          }
        : undefined,
    });
  }
}

export function performTabFollowWidth(operationState: PassOperationState, q: LayoutSeg): number {
  const { segAdvance, scale } = operationState;

  if ('isTab' in q) return q.measuredWidth || 0;
  if ('imagePath' in q) return q.widthPt * scale;
  if ('math' in q) return q.measuredWidth || 0;
  if ('lineBreak' in q) return 0;
  return segAdvance(q);
}

export function performDecimalAlignmentPoint(
  segments: readonly LayoutSeg[],
): Readonly<{ segmentIndex: number; charOffset: number }> | null {
  for (let segmentIndex = 0; segmentIndex < segments.length; segmentIndex += 1) {
    const segment = segments[segmentIndex]!;
    if (!('text' in segment)) continue;
    const separator = segment.text.indexOf('.');
    if (separator >= 0) return { segmentIndex, charOffset: separator };
  }

  let lastDigit: Readonly<{ segmentIndex: number; charOffset: number }> | null = null;
  let inFirstNumber = false;
  for (let segmentIndex = 0; segmentIndex < segments.length; segmentIndex += 1) {
    const segment = segments[segmentIndex]!;
    if (!('text' in segment)) {
      if (inFirstNumber) return lastDigit;
      continue;
    }
    let charOffset = 0;
    for (const scalar of segment.text) {
      charOffset += scalar.length;
      if (/\p{Decimal_Number}/u.test(scalar)) {
        inFirstNumber = true;
        lastDigit = { segmentIndex, charOffset };
      } else if (inFirstNumber) {
        return lastDigit;
      }
    }
  }
  return lastDigit;
}

export function performDecimalAlignmentPrefixWidth(
  operationState: PassOperationState,
  segments: readonly LayoutSeg[],
): number | undefined {
  const { strAdvance, tabFollowWidth, decimalAlignmentPoint } = operationState;

  const point = decimalAlignmentPoint(segments);
  if (!point) return undefined;
  let width = 0;
  for (let index = 0; index < point.segmentIndex; index += 1) {
    width += tabFollowWidth(segments[index]!);
  }
  const segment = segments[point.segmentIndex]!;
  if (!('text' in segment)) return width;
  return width + strAdvance(segment, segment.text.slice(0, point.charOffset));
}

export function performTabFollowingMetrics(operationState: PassOperationState): Readonly<{
  totalWidth: number;
  decimalPrefixWidth?: number;
}> {
  const { breakerState, tabFollowWidth, decimalAlignmentPrefixWidth } = operationState;

  const following: LayoutSeg[] = [];
  let totalWidth = 0;
  for (const q of breakerState.queue) {
    if ('isTab' in q || 'lineBreak' in q) break;
    following.push(q);
    totalWidth += tabFollowWidth(q);
  }
  const decimalPrefixWidth = decimalAlignmentPrefixWidth(following);
  return decimalPrefixWidth === undefined ? { totalWidth } : { totalWidth, decimalPrefixWidth };
}

export function performEmergencyTextSplit(
  operationState: PassOperationState,
  segment: LayoutTextSeg,
  available: number,
  forceAtLeastOne = true,
): number {
  const {
    prospectiveSnapAdvance,
    measurement,
    effectiveFontPx,
    setMeasureFont,
    strNaturalAdvance,
    strAdvance,
    ctx,
    scale,
    fontFamilyClasses,
    characterGrid,
    verticalGlyphMeasurement,
  } = operationState;

  const protectedOffsets = protectedNoBreakOffsets(segment);
  const graphemeOffsets = [0, ...graphemeClusterOffsets(segment.text), segment.text.length].filter(
    (offset, index, all) => all.indexOf(offset) === index,
  );
  let split = 0;
  if (available > 0) {
    const monotoneAllocation =
      charSpacingDeltaPx(segment, scale) >= 0 &&
      snapToCharsClass(segment, characterGrid) !== 'latin';
    if (monotoneAllocation) {
      setMeasureFont(
        buildFont(
          segment.bold,
          segment.italic,
          effectiveFontPx(segment),
          segment.fontFamily,
          fontFamilyClasses,
          segment.fontRoute,
        ),
      );
      measurement.withSegmentKerning(segment, () => {
        const fitted = fitCJKPrefix(
          ctx,
          segment.text,
          available,
          segmentCharacterGridDeltaPx(segment, characterGrid, scale),
          charScaleFactor(segment),
          charSpacingDeltaPx(segment, scale),
          segment.verticalRun === true,
          verticalGlyphMeasurement,
          (prefix) => strAdvance(segment, prefix),
        ).length;
        split =
          graphemeOffsets
            .filter((offset) => offset <= fitted && !protectedOffsets.has(offset))
            .at(-1) ?? 0;
      });
    } else {
      // Signed spacing and a Latin snap block can make prefix advances
      // non-monotone. Evaluate every legal retained candidate against the
      // exact prospective line block rather than binary-searching a
      // standalone approximation.
      for (const offset of graphemeOffsets) {
        if (offset <= 0 || protectedOffsets.has(offset)) continue;
        const natural = strNaturalAdvance(segment, segment.text.slice(0, offset));
        if (prospectiveSnapAdvance(segment, natural) <= available + 1e-9) split = offset;
      }
    }
  }
  if (split <= 0 && forceAtLeastOne) {
    split =
      graphemeOffsets.find((offset) => offset > 0 && !protectedOffsets.has(offset)) ??
      segment.text.length;
  }
  // Preserve the existing JLReq/Word line-end hanging rule after switching
  // the emergency splitter from code-point indexes to UTF-16 grapheme offsets.
  while (segment.text.startsWith('\u3000', split)) split += 1;
  return split;
}

export function performExternalLinkSyntaxSplit(
  operationState: PassOperationState,
  segment: LayoutTextSeg,
  available: number,
): number {
  const { prospectiveSnapAdvance, strNaturalAdvance } = operationState;

  if (!(available > 0) || !segment.externalLinkBreakOffsets?.length) return 0;
  let selected = 0;
  for (const offset of segment.externalLinkBreakOffsets) {
    if (offset <= 0 || offset >= segment.text.length) continue;
    const naturalAdvance = strNaturalAdvance(segment, segment.text.slice(0, offset));
    const prospectiveAdvance = prospectiveSnapAdvance(segment, naturalAdvance);
    if (prospectiveAdvance <= available + 1e-9) selected = offset;
  }
  return selected;
}

export function performQueueEmergencyTail(
  operationState: PassOperationState,
  segment: LayoutTextSeg,
  split: number,
): void {
  const { breakerState } = operationState;

  breakerState.queue.unshift({
    ...segment,
    ...RESET_SLICED_TEXT_MEASUREMENT,
    text: segment.text.slice(split),
    ...slicedTextMetadata(segment, split, segment.text.length),
    seaBreaks: rebaseSeaBreaks(segment.seaBreaks, split),
    measuredWidth: 0,
    // The emergency split itself is now the legal line boundary. A source-
    // boundary glue marker protects only the first retained prefix; carrying
    // it onto the tail would make the next line overflow again.
    joinPrev: undefined,
    hardJoinPrev: undefined,
    src: {
      segIndex: segment.src!.segIndex,
      charOffset: segment.src!.charOffset + split,
    },
  });
}

export function performRetractCurrentLineForLeadingKinsoku(
  operationState: PassOperationState,
  next: LayoutTextSeg,
): CrossRunKinsokuRetraction {
  const { breakerState, materializeLatinSpaceCompression, strAdvance, kinsoku } = operationState;
  return retractLeadingKinsoku(
    breakerState,
    kinsoku,
    materializeLatinSpaceCompression,
    strAdvance,
    next,
  );
}

export function performKeepLeadingKinsokuWithCurrentLine(
  operationState: PassOperationState,
  segment: LayoutTextSeg,
  h: number,
  asc: number,
  desc: number,
): boolean {
  const { breakerState, addToLine, strNaturalAdvance, queueEmergencyTail } = operationState;
  return keepLeadingKinsoku(
    breakerState,
    strNaturalAdvance,
    addToLine,
    queueEmergencyTail,
    segment,
    h,
    asc,
    desc,
  );
}
