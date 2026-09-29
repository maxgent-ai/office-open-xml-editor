import { describe, expect, it } from 'vitest';
import { excelDrawingMlLineRatios } from './office-auto-line.js';

const total = (r: ReturnType<typeof excelDrawingMlLineRatios>) => (r ? r.ascentRatio + r.descentRatio : NaN);

describe('excelDrawingMlLineRatios (#1604 Excel controls)', () => {
  it('uses the usWin box for bundled Far East faces (Yu Gothic 1.673 em, not hhea 1.433 em)', () => {
    const r = excelDrawingMlLineRatios({
      faceSource: 'office-bundle', unitsPerEm: 2048, hhea: [1802, -455, 1024], win: [2017, 619],
      farEastCodePage: true,
    })!;
    const box = (2017 + 619) / 2048;
    expect(r.ascentRatio).toBeCloseTo(2017 / 2048 + 0.15 * box, 10);
    expect(r.descentRatio).toBeCloseTo(619 / 2048 + 0.15 * box, 10);
  });

  it('uses usWin plus TEXTMETRIC external leading for other bundled faces', () => {
    // Baskerville Old Face: Excel 1.141 em (hhea 1.0 em).
    expect(total(excelDrawingMlLineRatios({
      faceSource: 'office-bundle', unitsPerEm: 2048, hhea: [1536, -512, 0], win: [1805, 531],
      farEastCodePage: false,
    }))).toBeCloseTo(2336 / 2048, 10);
    // A taller usWin box absorbs the hhea lineGap first.
    const clamp = excelDrawingMlLineRatios({
      faceSource: 'office-bundle', unitsPerEm: 1000, hhea: [900, -200, 150], win: [950, 250],
      farEastCodePage: false,
    })!;
    expect(clamp.ascentRatio).toBeCloseTo(1, 10);
  });

  it('uses typo metrics for a bundled face that sets USE_TYPO_METRICS (Gabriola 1.700 em)', () => {
    const r = excelDrawingMlLineRatios({
      faceSource: 'office-bundle', unitsPerEm: 4096, hhea: [2800, -1296, 2867], win: [4880, 2660],
      typoMetrics: [2800, -1296, 2867], farEastCodePage: false,
    })!;
    expect(r.ascentRatio).toBeCloseTo((2800 + 2867) / 4096, 10);
    expect(total(r)).toBeCloseTo((2800 + 1296 + 2867) / 4096, 10);
  });

  it('uses hhea for macOS system faces and leaves their Far East class unmeasured', () => {
    // Palatino: Excel 1.100 em (usWin 1.656 em).
    expect(total(excelDrawingMlLineRatios({
      faceSource: 'system', unitsPerEm: 2048, hhea: [1685, -568, 0], win: [2403, 989], farEastCodePage: false,
    }))).toBeCloseTo(2253 / 2048, 10);
    expect(excelDrawingMlLineRatios({
      faceSource: 'system', unitsPerEm: 1000, hhea: [880, -120, 0], win: [880, 120], farEastCodePage: true,
    })).toBeNull();
  });
});
