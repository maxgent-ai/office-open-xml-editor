import type { LayoutSeg } from './model.js';
import { wordCandidateFitWidthPx } from '../layout/line-compatibility.js';

/** Keep the fit decision on the same trailing-space policy as a line that will
 * later be justified. Manual and final lines retain their existing exception. */
export function justifiedCandidateFitWidth(
  widthPx: number,
  trailingSpacePx: number,
  next: LayoutSeg | undefined,
  state: Readonly<{
    isJustified: boolean;
    stretchLastLine: boolean;
    lineMaxWidth: number;
    lineXOffset: number;
    maxWidth: number;
  }>,
): number {
  const closesLogicalLine = next === undefined || 'lineBreak' in next;
  return wordCandidateFitWidthPx({
    widthPx,
    trailingSpacePx,
    lineWillJustify: state.isJustified && (!closesLogicalLine || state.stretchLastLine),
    wrapNarrowed: state.lineMaxWidth !== state.maxWidth || state.lineXOffset !== 0,
  });
}
