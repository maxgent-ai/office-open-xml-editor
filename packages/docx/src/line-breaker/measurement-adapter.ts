import { wordKerningApplies } from '../layout/line-compatibility.js';
import type { LayoutTextSeg } from '../line-layout.js';
import type { MeasurementTextContext, VerticalGlyphMeasurementService } from '../layout/measurement-capabilities.js';
import { calcEffectiveFontPx } from '../layout/text.js';
import { charScaleFactor } from './advance.js';
import { verticalRunInkExtra } from './vertical-text.js';

/** Owns the Canvas state used by one line-breaking pass. The state recorded here
 * describes only assignments made by this adapter; a future advance cache must
 * also account for the document-scoped shaping service and its feature state. */
export class LineMeasurementAdapter {
  private selectedFont: string | null = null;
  private readonly initialFont: string;
  private selectedKerning: CanvasFontKerning | null;
  readonly letterSpacing: string;

  constructor(
    private readonly context: MeasurementTextContext,
    private readonly scale: number,
    private readonly fontForSegment: (segment: LayoutTextSeg) => string,
    private readonly verticalGlyphMeasurement?: VerticalGlyphMeasurementService,
  ) {
    this.initialFont = context.font;
    this.selectedKerning = context.fontKerning;
    this.letterSpacing = context.letterSpacing;
  }

  setFont(font: string): void {
    if (font !== this.selectedFont) {
      this.context.font = font;
      this.selectedFont = font;
    }
  }

  selectSegmentFont(segment: LayoutTextSeg): void {
    this.setFont(this.fontForSegment(segment));
  }

  private selectSegmentKerning(segment: LayoutTextSeg): CanvasFontKerning | null {
    // ECMA-376 §17.3.2.19: an absent w:kern disables pair kerning even when
    // Canvas would otherwise choose its automatic kerning behavior.
    const previous = this.selectedKerning;
    // The acquisition request owns the mode-scoped decision; retained paint
    // consumes that same value. Standalone segments have no compatibility mode.
    const selected = (segment.textShapeRequest?.kerning ?? wordKerningApplies(segment.fontSize, segment.kerning))
      ? 'normal'
      : 'none';
    this.context.fontKerning = selected;
    this.selectedKerning = selected;
    return previous;
  }

  private restoreKerning(previous: CanvasFontKerning | null): void {
    if (previous != null) {
      this.context.fontKerning = previous;
      this.selectedKerning = previous;
    }
  }

  withSegmentKerning<T>(segment: LayoutTextSeg, measure: () => T): T {
    const previous = this.selectSegmentKerning(segment);
    try {
      return measure();
    } finally {
      this.restoreKerning(previous);
    }
  }

  measureSegment(segment: LayoutTextSeg, clusterGeometry: boolean | 'spaces' = false): TextMetrics {
    if (segment.textLayoutService && segment.textShapeRequest) {
      if (segment.textShapeRequest.text !== segment.text) {
        throw new Error('Segment measurement does not match its retained text range context');
      }
      const shaped = segment.textLayoutService.shape({
        ...segment.textShapeRequest,
        fontSizePt: calcEffectiveFontPx(segment, this.scale),
        measure: true,
        clusterGeometry,
      });
      if (clusterGeometry) {
        if (clusterGeometry === 'spaces') segment.shapedSpaceClusters = shaped.clusters;
        else segment.shapedClusters = shaped.clusters;
        segment.selectedFaceFontBox = {
          ascentPt: shaped.ascentPt,
          descentPt: shaped.descentPt,
        };
        segment.selectedFaceInkBounds = shaped.inkBounds ?? {
          xMinPt: 0,
          xMaxPt: shaped.advancePt,
          ascentPt: shaped.ascentPt,
          descentPt: shaped.descentPt,
        };
      }
      return {
        width: shaped.advancePt,
        actualBoundingBoxAscent: shaped.ascentPt,
        actualBoundingBoxDescent: shaped.descentPt,
        fontBoundingBoxAscent: shaped.ascentPt,
        fontBoundingBoxDescent: shaped.descentPt,
      } as TextMetrics;
    }
    this.selectSegmentFont(segment);
    const previous = this.selectSegmentKerning(segment);
    try {
      return this.context.measureText(segment.text);
    } finally {
      this.restoreKerning(previous);
    }
  }

  /** Browser pair-context repair at an actual ordinary word boundary. This
   * is native geometry, not an Office fitting allowance (§17.3.2.19).
   * Same-source, same-face horizontal non-complex text is measured as two
   * adjacent tokens and their concatenation; the difference belongs to the
   * next token's origin/advance. Each token is visited at most twice, without
   * a line-prefix cache. This does not promise arbitrary multi-token contextual
   * GSUB equivalence; RTL/complex and authored atomic units retain their own
   * shaping/placement contracts. Callers commit the result only on that line.
   * Library policy: a registered compound grapheme (semanticSlotSpans) is one
   * physical shape only as a single-grapheme request, so the joined probe
   * re-splits its §17.3.2.26 slots and would charge the detached mark's
   * advance to this boundary. Decline the repair there (no estimated pair
   * value), as the intrinsic-width merge does. A compound never ends with a
   * space, so only the right token can carry it. Unproven: whether native
   * one-string shaping kerns across such a boundary.
   */
  wordBoundaryAdvance(left: LayoutTextSeg | undefined, right: LayoutTextSeg): number {
    const l = left?.textShapeRequest;
    const r = right.textShapeRequest;
    const service = right.textLayoutService;
    if (!left || !l || !r || !service || service !== left.textLayoutService
      || !left.text.endsWith(' ') || right.text.startsWith(' ') || !right.text
      || right.semanticSlotSpans
      || left.metricOnly || right.metricOnly || left.ruby || right.ruby
      || left.fitTextRegionIndex !== undefined || right.fitTextRegionIndex !== undefined
      || left.verticalRun || right.verticalRun || left.rtl || right.rtl
      || l.complexScript || r.complexScript || l.kerning !== true
      || (left.script !== 'ascii' && left.script !== 'highAnsi')
      || left.script !== right.script
      || left.fontRoute?.fingerprint !== right.fontRoute?.fingerprint
      || calcEffectiveFontPx(left, this.scale) !== calcEffectiveFontPx(right, this.scale)
      || l.weight !== r.weight || l.style !== r.style || l.kerning !== r.kerning
      || charScaleFactor(left) !== charScaleFactor(right)
      || left.sourceRunIndex === undefined || left.sourceRunIndex !== right.sourceRunIndex
      || left.sourceTextSequence !== right.sourceTextSequence) return 0;
    const lc = l.substituteContext;
    const rc = r.substituteContext;
    if (!lc || !rc || lc.text !== rc.text || lc.offset + left.text.length !== rc.offset) return 0;
    const measure = (request: typeof r) => service.shape({ ...request,
      fontSizePt: calcEffectiveFontPx(right, this.scale), measure: true, clusterGeometry: false }).advancePt;
    const joined = { ...l, text: left.text + right.text };
    return (measure(joined) - measure(l) - measure(r)) * charScaleFactor(right);
  }

  measureRunText(segment: LayoutTextSeg, text: string): TextMetrics {
    this.selectSegmentFont(segment);
    const previous = this.selectSegmentKerning(segment);
    try {
      return this.context.measureText(text);
    } finally {
      this.restoreKerning(previous);
    }
  }

  measureCurrentText(text: string): TextMetrics {
    return this.context.measureText(text);
  }

  measureWithFont(font: string, text: string): TextMetrics {
    const previousFont = this.selectedFont ?? this.initialFont;
    this.setFont(font);
    try {
      return this.context.measureText(text);
    } finally {
      this.context.font = previousFont;
      this.selectedFont = previousFont;
    }
  }

  verticalInkExtra(segment: LayoutTextSeg, text: string): number {
    if (!segment.verticalRun) return 0;
    if (!this.verticalGlyphMeasurement) {
      throw new Error('Vertical glyph measurement capability is required for vertical text');
    }
    this.selectSegmentFont(segment);
    const previous = this.selectSegmentKerning(segment);
    try {
      return verticalRunInkExtra(text, true, this.verticalGlyphMeasurement);
    } finally {
      this.restoreKerning(previous);
    }
  }
}
