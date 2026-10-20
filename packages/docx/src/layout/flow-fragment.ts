import type { ParagraphLayout, TableLayout } from './types.js';

export type FlowFragment = ParagraphLayout | TableLayout;

export function fragmentLineAdvancesPt(fragment: ParagraphLayout): number {
  if (fragment.lines.length === 0) return fragment.paragraphMark?.bounds.heightPt ?? 0;
  let sum = 0;
  for (let index = 0; index < fragment.lines.length; index += 1) {
    const line = fragment.lines[index]!;
    if (index === 0) {
      sum += line.advancePt;
      continue;
    }
    const previous = fragment.lines[index - 1]!;
    sum += Math.max(0, line.bounds.yPt - previous.bounds.yPt - previous.advancePt)
      + line.advancePt;
  }
  return sum;
}

export function paragraphFragmentAdvancePt(fragment: ParagraphLayout): number {
  return fragment.advancePt;
}
