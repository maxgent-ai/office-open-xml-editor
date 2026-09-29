/**
 * `a:pPr@algn` (§20.1.10.59) in a physical horizontal text region. The
 * justifier owns slack distribution for `just`/`justLow`/`dist`/`thaiDist`;
 * this function supplies their unexpanded leading origin.
 */
export function drawingMlLineX(
  alignment: string | undefined,
  regionLeft: number,
  regionWidth: number,
  naturalWidth: number,
  rtl = false,
  /** PowerPoint preserves negative slack when text overflows a centred/right box. */
  allowOverflow = false,
): number {
  const slack = allowOverflow ? regionWidth - naturalWidth : Math.max(0, regionWidth - naturalWidth);
  if (alignment === 'ctr') return regionLeft + slack / 2;
  if (alignment === 'r' || (rtl && (alignment === 'l' || alignment == null))) {
    return regionLeft + slack;
  }
  return regionLeft;
}

/** `just`/`justLow` leave the final or manually terminated line ragged. */
export function drawingMlLineShouldJustify(
  alignment: string | undefined,
  lastLine: boolean,
  endsWithBreak: boolean,
): boolean {
  if (alignment === 'dist' || alignment === 'thaiDist') return true;
  if (alignment === 'just' || alignment === 'justLow') return !lastLine && !endsWithBreak;
  return false;
}
