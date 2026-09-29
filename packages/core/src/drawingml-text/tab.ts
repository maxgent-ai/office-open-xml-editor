/** Logical-order item used to resolve a DrawingML tab's advance. */
export interface DrawingMlTabItem {
  isTab: boolean;
  width: number;
}

/** An `a:pPr/a:tabLst/a:tab` stop in reading-frame pixels. */
export interface DrawingMlTabStop {
  pos: number;
  algn: string;
}

/**
 * Resolve `a:tab` gaps against the paragraph's explicit stops and `defTabSz`.
 * ECMA-376 §21.1.2.2.7 places both properties on `CT_TextParagraphProperties`.
 * The result is in logical reading order; the bidi pass reorders painted cells.
 * An infinite limit measures the natural extent for the break phase. A finite
 * limit clamps the paint placement to the trailing text edge.
 */
export function resolveDrawingMlTabWidths(
  items: readonly DrawingMlTabItem[],
  stops: readonly DrawingMlTabStop[],
  startPen: number,
  limit: number,
  noStopGap: number,
  defTabSz = 0,
): number[] {
  const widths = items.map((item) => item.width);
  const followW = (from: number): number => {
    let width = 0;
    for (let i = from; i < items.length && !items[i].isTab; i++) width += widths[i];
    return width;
  };

  let pen = startPen;
  for (let i = 0; i < items.length; i++) {
    if (!items[i].isTab) {
      pen += widths[i];
      continue;
    }

    let stop: DrawingMlTabStop | null = null;
    for (const candidate of stops) {
      if (candidate.pos > pen && (stop === null || candidate.pos < stop.pos)) {
        stop = candidate;
      }
    }
    if (stop === null) {
      if (defTabSz > 0) {
        stop = { pos: (Math.floor(pen / defTabSz) + 1) * defTabSz, algn: 'l' };
      } else {
        widths[i] = noStopGap;
        pen += noStopGap;
        continue;
      }
    }

    const followingWidth = followW(i + 1);
    const fraction = stop.algn === 'ctr'
      ? 0.5
      : stop.algn === 'r' || stop.algn === 'dec'
        ? 1
        : 0;
    let target = stop.pos - followingWidth * fraction;
    if (target + followingWidth > limit) target = limit - followingWidth;
    if (target < pen) target = pen;
    widths[i] = target - pen;
    pen = target;
  }
  return widths;
}
