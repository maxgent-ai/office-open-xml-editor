const FONT_PX = 20;

export interface FillCall {
  text: string;
  x: number;
  y: number;
  letterSpacing: string;
  font: string;
  fontKerning: CanvasFontKerning;
  // Naive accumulation of the raw translate() x arguments in effect at fill
  // time (save/restore-scoped, rotation ignored). Constant page-level offsets
  // cancel in the DELTAS between fills of the same line, which is all the
  // vertical assertions read (drawVerticalRun brackets each upright glyph in
  // save; translate(cellCenter, baseline); rotate; fillText; restore).
  translateX: number;
  translateY: number;
  scaleX: number;
  scaleY: number;
}

export function makeRecordingCanvas(options: Readonly<{
  followingGlyphLeftInsetPx?: number;
  closingPunctuationPairKerningPx?: number;
  spaceAdvancePx?: number;
  contextualShaping?: boolean;
  advanceFactor?: number;
}> = {}): { canvas: HTMLCanvasElement; fills: FillCall[] } {
  let font = `${FONT_PX}px serif`;
  let letterSpacing = '0px';
  let fontKerning: CanvasFontKerning = 'auto';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? String(FONT_PX));
  const fills: FillCall[] = [];
  let scaleX = 1;
  let scaleY = 1;
  let translateX = 0;
  let translateY = 0;
  const stack: { scaleX: number; scaleY: number; translateX: number; translateY: number }[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    get letterSpacing() { return letterSpacing; },
    set letterSpacing(v: string) { letterSpacing = v; },
    get fontKerning() { return fontKerning; },
    set fontKerning(value: CanvasFontKerning) { fontKerning = value; },
    measureText: (s: string) => {
      const p = px();
      const pairCount = s.match(/。）/gu)?.length ?? 0;
      const w = [...s].reduce(
        (sum, character) => sum + (options.contextualShaping && /[\p{M}\u00AD]/u.test(character)
          ? 0 : character === ' ' ? options.spaceAdvancePx ?? p : p * (options.advanceFactor ?? 1)),
        0,
      )
        - pairCount * (options.closingPunctuationPairKerningPx ?? 0)
        // A contextual metric makes loss of pair kerning/joining observable.
        // Synthetic metrics expose contextual shaping and fractional rounding.
        - (options.contextualShaping && fontKerning === 'normal'
          ? (s.match(/AV|To/gu)?.length ?? 0) * p / 5 : 0)
        - (options.contextualShaping ? (s.match(/سلام/gu)?.length ?? 0) * p * 0.6 : 0);
      return {
        width: w,
        actualBoundingBoxLeft:
          s === '次'
            ? -(options.followingGlyphLeftInsetPx ?? 0)
            : 0,
        // The punctuation fixture exposes a selected-glyph right ink edge at
        // half its full-width cell. Other text fills its complete advance.
        actualBoundingBoxRight:
          [...s].length === 1 && /[、．：しない）]/u.test(s) ? p / 2 : w,
        fontBoundingBoxAscent: p * 0.8,
        fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8,
        actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() { stack.push({ scaleX, scaleY, translateX, translateY }); },
    restore() { const s = stack.pop(); if (s) { scaleX = s.scaleX; scaleY = s.scaleY; translateX = s.translateX; translateY = s.translateY; } },
    beginPath() {}, closePath() {}, moveTo() {}, lineTo() {}, stroke() {}, fill() {},
    fillRect() {}, strokeRect() {}, clip() {}, rect() {}, setLineDash() {},
    drawImage() {}, clearRect() {}, arc() {}, quadraticCurveTo() {}, bezierCurveTo() {},
    createLinearGradient() { return { addColorStop() {} }; },
    scale(sx: number, sy?: number) { scaleX *= sx; scaleY *= sy ?? sx; },
    translate(tx: number, ty: number) { translateX += tx; translateY += ty; },
    rotate() {},
    setTransform() {},
    fillText(text: string, x: number, y: number) {
      fills.push({ text, x, y, font, fontKerning, letterSpacing, translateX, translateY, scaleX, scaleY });
    },
    strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, textBaseline: 'alphabetic' as CanvasTextBaseline,
    direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, fills };
}
