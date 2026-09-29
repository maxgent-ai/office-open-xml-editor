import { LineMeasurementAdapter } from './measurement-adapter.js';
import type { KinsokuRules } from '@silurus/ooxml-core';
import { containsSeaScript, seaMixedBreakOffsets } from '@silurus/ooxml-core';
import { type LayoutImageSeg, type LayoutLine, type LayoutMathSeg, type LayoutSeg, type LayoutTabSeg, type LayoutTextSeg, type LineBoundary, type WrapLayoutCtx } from './model.js';
import { protectedNoBreakOffsets, slicedTextMetadata } from './advance.js';
import { rebaseSeaBreaks } from './text-runs.js';
import { resolveFitTextSegments } from './segment-builder.js';

/** Prepare source-anchored break opportunities and the resumable queue.
 * SEA dictionary boundaries, protected ranges, paragraph-final hanging spaces,
 * and fitText allocation are resolved once per pass at the requested scale. */
export function prepareBreakQueue(
  segs: LayoutSeg[],
  startBoundary: LineBoundary | undefined,
  kinsoku: KinsokuRules,
  scale: number,
  measurement: LineMeasurementAdapter,
): LayoutSeg[] {
  const sourcedSegs = segs.map((seg, segIndex) => {
    seg.src = { segIndex, charOffset: 0 };
    // Issue #797 / #960 — attach the SEA (Thai/Lao/Khmer) break offsets ONCE per
    // segment (perf: never per line/char). Only for SEA text; non-SEA segments
    // keep `seaBreaks` absent so their wrap path is byte-identical. The set now
    // UNIONS the dictionary word boundaries (#797) with the no-space SEA↔non-SEA
    // script transitions and, for a mixed CJK+SEA run (a `<w:cs/>` run keeps CJK
    // in the same cs segment), the CJK per-character opportunities — so each
    // script keeps its own break rule inside one contiguous segment (#960). The
    // layout kinsoku set (§17.15.1.58–.60) drops positions that would orphan a
    // forbidden char at a line head/tail, replacing the CJK path's retract.
    if ('text' in seg && containsSeaScript(seg.text)) {
      const protectedOffsets = protectedNoBreakOffsets(seg);
      seg.seaBreaks = seaMixedBreakOffsets(seg.text, { cjk: true, kinsoku }).filter(
        (offset) => !protectedOffsets.has(offset),
      );
    }
    return seg;
  });
  let queue: LayoutSeg[];
  if (!startBoundary) {
    queue = sourcedSegs;
  } else if (startBoundary.segIndex >= sourcedSegs.length) {
    queue = [];
  } else {
    const first = sourcedSegs[startBoundary.segIndex];
    if (startBoundary.charOffset > 0) {
      if (!('text' in first) || startBoundary.charOffset > first.text.length) {
        queue = [];
      } else {
        const text = first.text.slice(startBoundary.charOffset);
        queue = text
          ? [
              {
                ...first,
                text,
                measuredWidth: 0,
                src: { ...startBoundary },
                // A retained resume boundary has already consumed the source
                // seam. Carrying either marker would invent new ownership at
                // the start of this suffix.
                joinPrev: undefined,
                hardJoinPrev: undefined,
                ...slicedTextMetadata(first, startBoundary.charOffset, first.text.length),
                // Rebase the SEA break offsets onto the resumed (sliced) text so
                // a paginated Thai paragraph still breaks at word boundaries.
                seaBreaks: rebaseSeaBreaks(first.seaBreaks, startBoundary.charOffset),
              },
              ...sourcedSegs.slice(startBoundary.segIndex + 1),
            ]
          : sourcedSegs.slice(startBoundary.segIndex + 1);
      }
    } else {
      queue = sourcedSegs.slice(startBoundary.segIndex);
    }
  }

  // Mark the paragraph-final U+3000 suffix once. A backwards pass avoids the
  // O(N²) suffix rescans that would result from queue.every/reduce per segment.
  let paragraphFinalIdeographicSpaceCount = 0;
  let paragraphFinalIdeographicSpaceTailStartIndex = -1;
  const markedParagraphFinalTail: Array<
    Readonly<{
      index: number;
      segment: LayoutTextSeg;
    }>
  > = [];
  for (let index = queue.length - 1; index >= 0; index -= 1) {
    const candidate = queue[index];
    if (!candidate || !('text' in candidate) || candidate.text.length === 0) break;
    // `fitText` (§17.3.2.14) and tate-chu-yoko (§17.3.2.10) are indivisible
    // layout cells. Ruby owns one base/guide pair. Paragraph-final whitespace
    // may affect how those cells measure, but must never split or clone them.
    if (
      candidate.fitTextRegionIndex !== undefined ||
      candidate.tateChuYoko === true ||
      candidate.ruby !== undefined
    ) {
      // buildSegments can split an atomic source run before its U+3000 tail
      // (ruby is retained only on the first emitted segment). Undo any markers
      // already assigned to trailing pieces from that same authored run.
      if (candidate.sourceRunIndex !== undefined) {
        for (
          let markedIndex = markedParagraphFinalTail.length - 1;
          markedIndex >= 0;
          markedIndex -= 1
        ) {
          const marked = markedParagraphFinalTail[markedIndex];
          if (marked.segment.sourceRunIndex !== candidate.sourceRunIndex) continue;
          marked.segment.paragraphFinalIdeographicSpaceTail = undefined;
          marked.segment.paragraphFinalIdeographicSpaceLocalCount = undefined;
          marked.segment.paragraphFinalIdeographicSpaceCount = undefined;
          marked.segment.paragraphFinalIdeographicSpaceTailStart = undefined;
          markedParagraphFinalTail.splice(markedIndex, 1);
        }
        paragraphFinalIdeographicSpaceTailStartIndex = markedParagraphFinalTail.at(-1)?.index ?? -1;
      }
      break;
    }
    const trailingSpaces = /^\u3000+$/u.test(candidate.text);
    const visibleWithTrailingSpaces = /[^\u3000]\u3000+$/u.test(candidate.text);
    if (!trailingSpaces && !visibleWithTrailingSpaces) break;
    const localTrailingCount = trailingSpaces
      ? [...candidate.text].length
      : [...candidate.text].reverse().findIndex((character) => character !== '\u3000');
    paragraphFinalIdeographicSpaceCount += localTrailingCount;
    candidate.paragraphFinalIdeographicSpaceTail = true;
    candidate.paragraphFinalIdeographicSpaceLocalCount = localTrailingCount;
    candidate.paragraphFinalIdeographicSpaceCount = paragraphFinalIdeographicSpaceCount;
    paragraphFinalIdeographicSpaceTailStartIndex = index;
    markedParagraphFinalTail.push({ index, segment: candidate });
    if (visibleWithTrailingSpaces) break;
  }
  if (paragraphFinalIdeographicSpaceTailStartIndex >= 0) {
    const start = queue[paragraphFinalIdeographicSpaceTailStartIndex];
    if (start && 'text' in start) start.paragraphFinalIdeographicSpaceTailStart = true;
  }

  // Resolve §17.3.2.14 from RAW natural advances at this exact layout scale.
  // The resulting per-gap is folded into segAdvanceWidth below, so the line
  // breaker and paint pen use one width authority. Cached w:spacing is ignored.
  // #1014 — the natural width includes the vo=Tr ink deficit so the resolved gap
  // (target − natural)/n, plus the ink-grown cell the paint draws, still sums to
  // the fitText target (measure == paint); 0 for non-under-reporting runs.
  resolveFitTextSegments(
    queue.filter((seg): seg is LayoutTextSeg => 'text' in seg),
    scale,
    (segment) =>
      measurement.measureSegment(segment).width +
      measurement.verticalInkExtra(segment, segment.text),
  );

  return queue;
}


export type SnapBlockState = {
    kind: 'latin' | 'complexScript';
    first: LayoutTextSeg;
    last: LayoutTextSeg;
    naturalWidthPx: number;
    allocatedWidthPx: number;
  };


export function createLineBreakerState(maxWidth: number, wrapCtx?: WrapLayoutCtx) {
  return {
    lines: [] as LayoutLine[],
    currentLine: [] as (LayoutTextSeg | LayoutImageSeg | LayoutMathSeg | LayoutTabSeg)[],
    currentWidth: 0,
    latinLineFace: undefined as LayoutTextSeg | undefined,
    latinLineHomogeneous: true,
    latinLineGaps: [] as LayoutTextSeg[],
    latinUniformGapCapacity: undefined as number | undefined,
    latinAppliedGapCount: 0,
    latinAppliedPerGap: 0,
    snapBlock: null as SnapBlockState | null,
    lineHeight: 0,
    lineAscent: 0,
    lineDescent: 0,
    lineIntendedSingle: 0,
    lineHasInlinePicture: false,
    linePictureMarkSingle: 0,
    lineGridCountSingle: 0,
    lineVisibleAscent: 0,
    lineVisibleDescent: 0,
    lineVisibleIntendedSingle: 0,
    lineHasVisibleMetrics: false,
    isFirst: true,
    lineMaxWidth: maxWidth,
    lineXOffset: 0,
    currentLineTopY: wrapCtx?.startPageY ?? 0,
    lineHasRuby: false,
    lineEastAsian: false,
    queue: [] as LayoutSeg[],
    trailingBreakFontSize: null as number | null,
  };
}
