export interface FitTextRun {
  /** §17.3.2.14 w:val — target region width in TWIPS. undefined ⇒ not a fitText run. */
  fitTextValTwips?: number;
  /** §17.3.2.14 w:id — links CONSECUTIVE fitText runs into ONE region.
   *  undefined ⇒ id 0 (WORD_FIT_TEXT_OMITTED_ID_LINKING). */
  fitTextId?: number | string;
  /** Code-point count of the run. */
  charCount: number;
  /** Natural glyph-advance SUM of the run in px (at the layout scale), BEFORE the
   *  §17.3.2.43 w:w scale and WITHOUT any §17.3.2.35 w:spacing pitch. */
  naturalWidthPx: number;
  /** §17.3.2.43 w:w horizontal glyph scale as a FRACTION (0.66 = 66%). undefined ⇒ 1. */
  charScale?: number;
}

export interface FitTextRegion {
  start: number;
  end: number;
  targetPx: number;
  naturalPx: number;
  charCount: number;
  perGapPx: number;
  trailingPadPx: number;
}

/** The region id of a fitText run (`WORD_FIT_TEXT_OMITTED_ID_LINKING`): an
 *  omitted w:id is id 0. Integer ids compare by value (the XSD integer is
 *  arbitrary precision, so BigInt, not Number). */
function fitTextRegionId(id: number | string | undefined): string {
  if (id === undefined) return '0';
  const text = String(id).trim();
  return /^[+-]?\d+$/.test(text) ? BigInt(text).toString() : text;
}

/** ECMA-376 §17.3.2.14. `scale` is px-per-pt. */
export function groupFitTextRegions(runs: FitTextRun[], scale: number): FitTextRegion[] {
  const regions: FitTextRegion[] = [];

  for (let start = 0; start < runs.length; ) {
    const first = runs[start];
    if (first.fitTextValTwips === undefined) {
      start += 1;
      continue;
    }

    let end = start + 1;
    const id = fitTextRegionId(first.fitTextId);
    while (
      end < runs.length &&
      runs[end].fitTextValTwips !== undefined &&
      fitTextRegionId(runs[end].fitTextId) === id
    ) {
      end += 1;
    }

    let naturalPx = 0;
    let charCount = 0;
    for (let index = start; index < end; index += 1) {
      const run = runs[index];
      naturalPx += run.naturalWidthPx * (run.charScale ?? 1);
      charCount += run.charCount;
    }

    const targetPx = (first.fitTextValTwips / 20) * scale;
    // ECMA-376 §17.3.2.14 requires the region to occupy exactly w:val and
    // describes compression as “decreasing the size of each character”.
    // `word-fit-text-inter-character-expansion`: expand a multi-character region
    // through inter-character gaps. Compression geometry remains unsupported;
    // until separately adjudicated, keep each glyph at natural width and use the
    // same gap formula. A one-character region has no gap, so its residual is
    // trailing cell padding rather than glyph scaling.
    const perGapPx = charCount > 1 ? (targetPx - naturalPx) / (charCount - 1) : 0;
    const trailingPadPx = targetPx - naturalPx - Math.max(0, charCount - 1) * perGapPx;

    regions.push({ start, end, targetPx, naturalPx, charCount, perGapPx, trailingPadPx });
    start = end;
  }

  return regions;
}
