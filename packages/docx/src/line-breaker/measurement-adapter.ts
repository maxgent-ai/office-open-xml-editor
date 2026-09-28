import type { LayoutTextSeg } from '../line-layout.js';
import type { MeasurementTextContext, VerticalGlyphMeasurementService } from '../layout/measurement-capabilities.js';
import { calcEffectiveFontPx } from '../layout/text.js';
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
    const selected = segment.kerning != null && segment.fontSize >= segment.kerning
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

  measureSegment(segment: LayoutTextSeg, clusterGeometry = false): TextMetrics {
    if (segment.textLayoutService && segment.textShapeRequest) {
      const shaped = segment.textLayoutService.shape({
        ...segment.textShapeRequest,
        text: segment.text,
        fontSizePt: calcEffectiveFontPx(segment, this.scale),
        measure: true,
        clusterGeometry,
      });
      if (clusterGeometry) {
        segment.shapedClusters = shaped.clusters;
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
