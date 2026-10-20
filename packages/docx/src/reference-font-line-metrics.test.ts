import { describe, expect, it, vi } from 'vitest';
import { referenceFontLineMetrics } from './reference-font-line-metrics.js';

describe('DOCX reference font vertical metrics', () => {
  it('projects Calibri regular hhea geometry for automatic line spacing', () => {
    const metric = referenceFontLineMetrics('Calibri', 400, 'normal');
    expect(metric).toEqual({
      lineHeightRatio: 2500 / 2048,
      farEastCodePage: false,
      designAscentRatio: 1950 / 2048,
      designDescentRatio: 550 / 2048,
    });
    expect(11 * metric!.lineHeightRatio * (259 / 240))
      .toBeCloseTo(14.490763346354166, 12);
  });

  it.each([
    ['MacIntel', 'darwin'],
    ['Win32', 'win32'],
    [undefined, 'darwin'],
    [undefined, 'linux'],
  ])('keeps the same unmeasurable-host catalog fallback with navigator.platform=%s and process.platform=%s', async (
    browserPlatform, nodePlatform,
  ) => {
    vi.stubGlobal('navigator', browserPlatform === undefined ? undefined : { platform: browserPlatform });
    vi.stubGlobal('process', { ...process, platform: nodePlatform });
    vi.resetModules();
    try {
      const { referenceFontLineMetrics: vertical, referenceFontAverageWidthRatio: averageWidth }
        = await import('./reference-font-line-metrics.js');
      // Both catalogs contain these tuples with different metrics. Reloading
      // under each host catches source selection performed at module startup.
      expect(vertical('Times New Roman')?.lineHeightRatio).toBe(2268 / 2048);
      expect(vertical('Symbol')).toEqual({
        lineHeightRatio: (2059 + 450) / 2048,
        farEastCodePage: false,
        designAscentRatio: 2059 / 2048,
        designDescentRatio: 450 / 2048,
      });
      expect(averageWidth('Symbol')).toBe(1229 / 2048);
    } finally {
      vi.unstubAllGlobals();
      vi.resetModules();
    }
  });

  it('uses the Office Far-East allocation for Meiryo even on a Latin line', () => {
    const metric = referenceFontLineMetrics('Meiryo', 400, 'normal');
    const box = 2171 + 901;
    expect(metric?.lineHeightRatio).toBeCloseTo(1.3 * box / 2048, 12);
    expect(metric?.designAscentRatio).toBeCloseTo((2171 + 0.15 * box) / 2048, 12);
    expect(metric?.designDescentRatio).toBeCloseTo((901 + 0.15 * box) / 2048, 12);
  });

  it('projects Yu Mincho from its own usWin sides instead of hhea', () => {
    const metric = referenceFontLineMetrics('Yu Mincho', 400, 'normal');
    expect(metric?.lineHeightRatio).toBeCloseTo(1.3 * (2038 + 598) / 2048, 12);
    expect(metric?.designAscentRatio).toBeCloseTo((2038 + .15 * 2636) / 2048, 12);
    expect(metric?.designDescentRatio).toBeCloseTo((598 + .15 * 2636) / 2048, 12);
  });

  it('retains signed hhea leading, including a macOS-only fallback with a negative gap', () => {
    const office = referenceFontLineMetrics('Times New Roman', 400, 'normal');
    expect(office?.lineHeightRatio).toBe((1825 + 443) / 2048);
    expect(office?.designAscentRatio).toBe(1825 / 2048);
    const negativeGap = referenceFontLineMetrics('Tamil Sangam MN', 400, 'normal');
    expect(negativeGap?.lineHeightRatio).toBe((1550 + 717 - 210) / 2048);
    expect(negativeGap?.designAscentRatio).toBe((1550 - 210) / 2048);
  });

  it('declines an unknown family and an unavailable style tuple', () => {
    expect(referenceFontLineMetrics('Not In Reference Catalog')).toBeUndefined();
    expect(referenceFontLineMetrics('Calibri', 900, 'italic')).toBeUndefined();
    // OS/2 code-page data is absent; a listed family alone cannot choose the
    // ordinary or Far-East Word allocation class.
    expect(referenceFontLineMetrics('AppleGothic', 400, 'normal')).toBeUndefined();
  });
});
