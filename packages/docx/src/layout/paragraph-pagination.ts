import { adjustForWidowOrphan, selectLargestFittingEnd } from '../line-fit-policy.js';
import type { LineBoundary } from '../line-layout.js';
import {
  wordFinalParagraphAdmissionExtentPt,
  wordVerticalRlFinalLineAdmissionExtentPt,
} from './body-pagination-compatibility.js';
import { sliceParagraphLayout } from './paragraph.js';
import type { ParagraphLayout, WritingMode } from './types.js';

export interface ParagraphFragmentCursor {
  readonly boundary: LineBoundary | null;
  readonly sourceRangeStart?: number;
  readonly uniformRubyAdvancePt?: number;
}

export interface ParagraphFragmentSelection {
  readonly fragment: ParagraphLayout | null;
  readonly nextCursor: ParagraphFragmentCursor | null;
  readonly requiresFreshFlowRegion: boolean;
  readonly additionalReservePt: number;
  /** Flow charge admitted to this region; retained line/ink geometry can be taller. */
  readonly admittedBlockExtentPt: number;
}

export type ParagraphFragmentation =
  | Readonly<{ kind: 'indivisible' }>
  | Readonly<{
      kind: 'splittable';
      /** Source cursor after every retained visual line, including the final line. */
      lineEndBoundaries: readonly LineBoundary[];
    }>;

function compareLineBoundaries(left: LineBoundary, right: LineBoundary): number {
  return left.segIndex - right.segIndex || left.charOffset - right.charOffset;
}

export function selectParagraphFragment(
  acquired: ParagraphLayout,
  cursor: ParagraphFragmentCursor,
  fragmentation: ParagraphFragmentation,
  availableBlockExtentPt: number,
  freshFlowRegionBlockExtentPt: number,
  canRelocate: boolean,
  policy: Readonly<{
    keepLines: boolean;
    widowControl: boolean;
    /** Authored §17.3.1.33 trailing whitespace; final-fragment fit is governed
     * by WORD_TRAILING_SPACE_AFTER_FIT_ADMISSION. */
    authoredSpaceAfterPt?: number;
    /** Owning section-region flow axis. Vertical final-line admission is
     * governed by WORD_VERTICAL_RL_FINAL_LINE_BASELINE_ADMISSION. */
    writingMode?: WritingMode;
  }>,
  additionalReserveFor?: (fragment: ParagraphLayout) => number,
  uniformRubyAdvancePt?: number,
  additionalReserveFits?: (reservePt: number) => boolean,
): ParagraphFragmentSelection {
  if (![availableBlockExtentPt, freshFlowRegionBlockExtentPt].every(
    (value) => Number.isFinite(value) && value >= 0,
  )) throw new RangeError('Paragraph fragment extents must be finite and non-negative');
  if (
    fragmentation.kind === 'splittable'
    && fragmentation.lineEndBoundaries.length !== acquired.lines.length
  ) {
    throw new RangeError('Splittable paragraph source boundaries must align with retained lines');
  }
  if (fragmentation.kind === 'indivisible' && cursor.boundary !== null) {
    throw new RangeError('Indivisible paragraph cannot carry a continuation boundary');
  }
  const authoredSpaceAfterPt = policy.authoredSpaceAfterPt ?? 0;
  if (!Number.isFinite(authoredSpaceAfterPt) || authoredSpaceAfterPt < 0) {
    throw new RangeError('Authored paragraph spaceAfter must be finite and non-negative');
  }
  const total = acquired.lines.length;
  const slice = (end: number) => sliceParagraphLayout(acquired, {
    lineStart: 0,
    lineEnd: end,
    continuesFromPrevious: cursor.boundary !== null,
    continuesOnNext: end < total,
  });
  const reserveFor = (fragment: ParagraphLayout): number => {
    const reserve = additionalReserveFor?.(fragment) ?? 0;
    if (!Number.isFinite(reserve) || reserve < 0) {
      throw new RangeError('Paragraph page-local reserve must be finite and non-negative');
    }
    return reserve;
  };
  const reserveFits = (reservePt: number): boolean => additionalReserveFits?.(reservePt) ?? true;
  const admissionExtent = (fragment: ParagraphLayout, completesParagraph: boolean): number => {
    if (!completesParagraph) return fragment.advancePt;
    // Keep retained advance authoritative for placement and paint. Only the
    // page/column fit comparison ignores authored trailing whitespace; any
    // retained trailing extent beyond it (for example a bottom border) remains.
    const logicalLineBoxExtentPt = wordFinalParagraphAdmissionExtentPt({
      advancePt: fragment.advancePt,
      retainedSpaceAfterPt: fragment.spacing.afterPt,
      authoredSpaceAfterPt,
    });
    return wordVerticalRlFinalLineAdmissionExtentPt({
      paragraph: fragment,
      writingMode: policy.writingMode ?? 'horizontal-tb',
      logicalLineBoxExtentPt,
      availableBlockExtentPt,
    });
  };
  if (fragmentation.kind === 'indivisible') {
    const completeReserve = reserveFor(acquired);
    const completeExtentPt = admissionExtent(acquired, true);
    if (canRelocate && (
      completeExtentPt + completeReserve > availableBlockExtentPt
      || !reserveFits(completeReserve)
    ) && completeExtentPt + completeReserve <= freshFlowRegionBlockExtentPt) {
      return {
        fragment: null, nextCursor: cursor,
        requiresFreshFlowRegion: true, additionalReservePt: 0, admittedBlockExtentPt: 0,
      };
    }
    return {
      fragment: acquired,
      nextCursor: null,
      requiresFreshFlowRegion: false,
      additionalReservePt: completeReserve,
      admittedBlockExtentPt: Math.min(acquired.advancePt, availableBlockExtentPt),
    };
  }
  if (total === 0) {
    const reserve = reserveFor(acquired);
    const completeExtentPt = admissionExtent(acquired, true);
    if (canRelocate && (
      completeExtentPt + reserve > availableBlockExtentPt
      || !reserveFits(reserve)
    )
      && completeExtentPt + reserve <= freshFlowRegionBlockExtentPt) {
      return {
        fragment: null, nextCursor: cursor,
        requiresFreshFlowRegion: true, additionalReservePt: 0, admittedBlockExtentPt: 0,
      };
    }
    return {
      fragment: acquired, nextCursor: null,
      requiresFreshFlowRegion: false, additionalReservePt: reserve,
      admittedBlockExtentPt: Math.min(acquired.advancePt, availableBlockExtentPt),
    };
  }
  const completeReserve = reserveFor(acquired);
  const completeExtentPt = admissionExtent(acquired, true);
  if (cursor.boundary === null && policy.keepLines && canRelocate
    && (
      completeExtentPt + completeReserve > availableBlockExtentPt
      || !reserveFits(completeReserve)
    )
    && completeExtentPt + completeReserve <= freshFlowRegionBlockExtentPt) {
    return {
      fragment: null, nextCursor: cursor,
      requiresFreshFlowRegion: true, additionalReservePt: 0, admittedBlockExtentPt: 0,
    };
  }
  let end = selectLargestFittingEnd(
    0,
    total,
    availableBlockExtentPt,
    (lineEnd) => (() => {
      const candidate = slice(lineEnd);
      const reserve = reserveFor(candidate);
      return reserveFits(reserve)
        ? admissionExtent(candidate, lineEnd === total) + reserve
        : availableBlockExtentPt + 1;
    })(),
  ).end;
  if (end === 0) {
    if (canRelocate) return {
      fragment: null, nextCursor: cursor,
      requiresFreshFlowRegion: true, additionalReservePt: 0, admittedBlockExtentPt: 0,
    };
    end = 1;
  }
  for (;;) {
    const widow = adjustForWidowOrphan({
      widowControl: policy.widowControl,
      start: 0,
      end,
      totalLines: total,
      canRelocate,
    });
    if (widow.kind === 'relocate') {
      return {
        fragment: null, nextCursor: cursor,
        requiresFreshFlowRegion: true, additionalReservePt: 0, admittedBlockExtentPt: 0,
      };
    }
    if (widow.kind !== 'dropLastLine') break;
    end -= 1;
  }
  const fragment = slice(end);
  const nextBoundary = end < total ? fragmentation.lineEndBoundaries[end - 1]! : null;
  if (
    nextBoundary !== null
    && cursor.boundary !== null
    && compareLineBoundaries(nextBoundary, cursor.boundary) <= 0
  ) {
    throw new Error('Paragraph continuation source boundary did not advance');
  }
  return {
    fragment,
    nextCursor: nextBoundary === null ? null : Object.freeze({
      boundary: nextBoundary,
      sourceRangeStart: fragment.lines.at(-1)!.range.end,
      ...(uniformRubyAdvancePt === undefined ? {} : { uniformRubyAdvancePt }),
    }),
    requiresFreshFlowRegion: false,
    additionalReservePt: reserveFor(fragment),
    admittedBlockExtentPt: Math.min(fragment.advancePt, availableBlockExtentPt),
  };
}
