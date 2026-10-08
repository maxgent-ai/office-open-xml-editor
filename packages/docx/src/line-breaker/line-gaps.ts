import { enumerateGaps, type GapScanState, type LineGap, type DistributeOptions } from '@silurus/ooxml-core';
import type { LayoutSeg } from './model.js';

/** Width authority is the breaker's advance or retained layout geometry, never
 * a font-name approximation. Opaque cells include tabs, inline objects and ruby;
 * zero-advance metric hosts have empty text and do not interrupt adjacency. */
export interface GapSegment {
  text?: string;
  widthPx: number;
  spacePx: number;
  /** Contextual per-space advances, indexed in code points. */
  spaceWidths?: ReadonlyMap<number, number>;
}

export const textCellGapOptions: DistributeOptions = {
  textCellSpaceGaps: true,
  atomicSpaceGaps: true,
  isWhitespace: cp => cp === 0x20,
  isGapChar: () => false,
  lastDrawnSi: Infinity,
};

/** Complete opportunities retained once when the accepted line closes. The
 * streaming prefix remains constant-size during candidate comparisons. */
export interface LineGapPlan extends LineGapModel {
  expansionGaps: LineGap[];
}

export interface LineGapModel {
  gaps: LineGap[];
  S: number;
  visibleWidthPx: number;
  lineEndSeparatorPx: number;
  naturalWidthPx: number;
  segmentCount: number;
  scan: GapScanState;
}

/** WORD_JUSTIFIED_INTERWORD_COMPRESSION: the same enumerator owns fit and
 * positive/negative justification. Consecutive U+0020 all participate only
 * between text cells; NBSP, tab/object/ruby adjacency and line edges stay fixed.
 * Prefix accumulation retains constant-size state, so candidate lookahead and
 * committed appends are linear even with long fragmented words/space runs. */
export function lineGapModel(
  segments: readonly GapSegment[],
  prefix?: Readonly<LineGapModel>,
  collect = true,
): LineGapModel {
  const offset = prefix?.segmentCount ?? 0;
  const enumeration = enumerateGaps(segments, {
    ...textCellGapOptions,
    segmentOffset: offset,
    prefix: prefix?.scan,
    collect,
    gapWidth: (si, cpOffset) => segments[si - offset].spaceWidths?.get(cpOffset) ?? segments[si - offset].spacePx,
  });
  let naturalWidthPx = prefix?.naturalWidthPx ?? 0;
  let lineEndSeparatorPx = prefix?.lineEndSeparatorPx ?? 0;
  for (const segment of segments) {
    naturalWidthPx += segment.widthPx;
    if (segment.text === '') continue;
    if (segment.text !== undefined && /^ +$/u.test(segment.text)) {
      lineEndSeparatorPx += segment.widthPx;
    } else {
      let count = 0;
      const text = segment.text ?? '';
      for (let offset = text.length - 1; offset >= 0 && text[offset] === ' '; offset--) count += 1;
      lineEndSeparatorPx = 0;
      const length = [...text].length;
      for (let off = length - count; off < length; off++) {
        lineEndSeparatorPx += segment.spaceWidths?.get(off) ?? segment.spacePx;
      }
    }
  }
  return {
    gaps: enumeration.gaps,
    S: enumeration.state.naturalGapSum,
    naturalWidthPx,
    visibleWidthPx: naturalWidthPx - lineEndSeparatorPx,
    lineEndSeparatorPx,
    segmentCount: offset + segments.length,
    scan: enumeration.state,
  };
}

/** A complete next text unit, including its separator runs. Opaque cells are
 * indivisible units of their own. No following-word eligibility lookahead. */
export function candidateUnit(first: LayoutSeg, queue: Iterable<LayoutSeg>): LayoutSeg[] {
  const segments = [first];
  if (!('text' in first) || first.ruby || first.fitTextRegionIndex !== undefined) return segments;
  let hasSeparator = first.text.endsWith(' ');
  for (const next of queue) {
    if (!('text' in next) || next.ruby || next.fitTextRegionIndex !== undefined
      || (hasSeparator && /[^ ]/u.test(next.text) && !next.joinPrev && !next.hardJoinPrev)) break;
    segments.push(next);
    if (next.text.length > 0) hasSeparator = next.text.endsWith(' ');
  }
  return segments;
}
