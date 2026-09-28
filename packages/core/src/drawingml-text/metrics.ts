/** `a:bodyPr` text box after its four independent EMU insets. */
export interface DrawingMlInsets {
  lIns: number;
  rIns: number;
  tIns: number;
  bIns: number;
}

export interface DrawingMlTextRect {
  left: number;
  top: number;
  width: number;
  height: number;
}

/** ECMA-376 §21.1.2.1.1: all four `bodyPr` insets affect the text rectangle. */
export function drawingMlTextRect(
  width: number,
  height: number,
  insets: DrawingMlInsets,
  pxPerEmu: number,
): DrawingMlTextRect {
  const left = insets.lIns * pxPerEmu;
  const right = insets.rIns * pxPerEmu;
  const top = insets.tIns * pxPerEmu;
  const bottom = insets.bIns * pxPerEmu;
  return {
    left,
    top,
    width: Math.max(0, width - left - right),
    height: Math.max(0, height - top - bottom),
  };
}

export type DrawingMlLineSpacing =
  | { type: 'pct'; val: number }
  | { type: 'pts'; val: number }
  | null
  | undefined;

/**
 * `a:lnSpc` (§21.1.2.2.5) and `a:normAutofit@lnSpcReduction`
 * (§21.1.2.1.3). Percentage spacing multiplies the natural line box;
 * point spacing replaces it. The stored normal-autofit reduction applies
 * only to percentage spacing (including the implicit 100%).
 */
export function drawingMlLineHeight(
  naturalHeight: number,
  spacing: DrawingMlLineSpacing,
  pxPerPt: number,
  reduction = 0,
): number {
  const base = spacing?.type === 'pts'
    ? spacing.val * pxPerPt
    : naturalHeight * (spacing?.type === 'pct' ? spacing.val / 100000 : 1);
  return spacing?.type === 'pts' ? base : base * (1 - reduction);
}

/** ECMA-376 §20.1.7.2 `a:bodyPr@anchor` for the line block. */
export function drawingMlBlockTop(
  anchor: string | undefined,
  rect: DrawingMlTextRect,
  blockHeight: number,
): number {
  if (anchor === 'ctr') return rect.top + (rect.height - blockHeight) / 2;
  if (anchor === 'b') return rect.top + Math.max(0, rect.height - blockHeight);
  return rect.top;
}
