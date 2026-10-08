import { describe, expect, it } from 'vitest';
import type { LayoutTextSeg } from '../line-layout.js';
import type { MeasurementTextContext } from '../layout/measurement-capabilities.js';
import { LineMeasurementAdapter } from './measurement-adapter.js';

describe('LineMeasurementAdapter', () => {
  it('measures interleaved faces with each run kerning rule and restores the caller state', () => {
    const context: MeasurementTextContext = {
      font: 'initial',
      fontKerning: 'auto',
      letterSpacing: '0px',
      measureText(text) {
        return {
          width: text.length * (this.font === 'wide' ? 20 : 10)
            + (this.fontKerning === 'normal' ? 5 : 0),
        } as TextMetrics;
      },
    };
    const adapter = new LineMeasurementAdapter(
      context,
      1,
      (segment) => segment.fontFamily ?? 'initial',
    );
    const narrow = {
      text: 'AB', fontFamily: 'narrow', fontSize: 10, kerning: 12,
    } as LayoutTextSeg;
    const wide = {
      text: 'AB', fontFamily: 'wide', fontSize: 10, kerning: 8,
    } as LayoutTextSeg;

    expect(adapter.measureSegment(narrow).width).toBe(20);
    expect(adapter.measureSegment(wide).width).toBe(45);
    expect(adapter.measureRunText(narrow, 'A').width).toBe(10);
    expect(adapter.measureWithFont('wide', 'A').width).toBe(20);
    expect(context.font).toBe('narrow');
    expect(context.fontKerning).toBe('auto');
  });
});
