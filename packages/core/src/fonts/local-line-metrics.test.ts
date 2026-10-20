import { describe, expect, it, vi } from 'vitest';
import { localFontMetricCoversText, measureLocalFontNormalLine } from './local-line-metrics.js';
import type { ResolvedFontMetric } from './resource-metrics.js';
import type { CanvasFontRoute } from './canvas-route.js';

// Boxes are given at a 2048px probe and scale with the requested size, like
// an engine without a size ceiling; `clampPx` emulates an engine that clamps.
function environment(height = 2355, clampPx = Infinity) {
  let reads = 0;
  let removed = 0;
  const px = (font: string) => Math.min(Number(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? 0), clampPx);
  const probe = {
    style: { cssText: '' }, textContent: '',
    get offsetHeight() { reads++; return Math.round(height * px(probe.style.cssText) / 2048); },
  };
  const host = {
    style: { cssText: '' },
    attachShadow: () => ({ append() {} }),
    remove() { removed++; },
  };
  const fonts = {} as FontFaceSet;
  // `family` emulates the computed font-family; an application stylesheet
  // overriding the strut would change it.
  const computed = { family: undefined as string | undefined };
  const document = {
    fonts,
    body: { append() {} },
    createElement: (tag: string) => tag === 'div' ? host : probe,
    defaultView: {
      getComputedStyle: () => ({
        fontFamily: computed.family ?? /px ("[^"]*")/.exec(probe.style.cssText)?.[1] ?? '',
      }),
    },
  } as unknown as Document;
  const context = {
    font: '12px serif',
    measureText() {
      const scale = px(this.font) / 2048;
      return { fontBoundingBoxAscent: 1825 * scale, fontBoundingBoxDescent: 443 * scale } as TextMetrics;
    },
  };
  return { document, fonts, context, probe, computed, reads: () => reads, removed: () => removed };
}

describe('prepared local-face normal line measurement', () => {
  it('retains the host line gap once per loaded resource, using its exact paint alias', () => {
    const env = environment();
    const face = { family: 'isolated-face', weight: '400', style: 'normal', status: 'loaded' } as FontFace;
    const first = measureLocalFontNormalLine(face, env.fonts, env.context, 2048, env.document);
    expect(first).toEqual({ ascentRatio: 1825 / 2048, descentRatio: 443 / 2048, lineGapRatio: 87 / 2048 });
    expect(measureLocalFontNormalLine(face, env.fonts, env.context, 2048, env.document)).toBe(first);
    // One linearity check: the requested em and half of it.
    expect(env.reads()).toBe(2);
    expect(env.removed()).toBe(2);
    expect(env.probe.style.cssText).toContain('"isolated-face"');
    expect(env.context.font).toBe('12px serif');
    // An equal alias in another registration must not reuse this resource's gap.
    const other = environment(2268);
    expect(measureLocalFontNormalLine({ ...face } as FontFace, other.fonts, other.context, 2048, other.document)?.lineGapRatio).toBe(0);
  });

  it('normalizes by a linearly measured em when the engine clamps large sizes', () => {
    const face = { family: 'isolated-face', weight: '400', style: 'normal', status: 'loaded' } as FontFace;
    // A 2000px ceiling makes the 2048px box non-linear against 1024px; the
    // 1024px box is linear against 512px and supplies the host ratios.
    const env = environment(2355, 2000);
    const metric = measureLocalFontNormalLine(face, env.fonts, env.context, 2048, env.document);
    expect(metric?.ascentRatio).toBeCloseTo(1825 / 2048, 6);
    expect(metric?.descentRatio).toBeCloseTo(443 / 2048, 6);
    expect(metric?.lineGapRatio).toBeCloseTo((Math.round(2355 / 2) - 1825 / 2 - 443 / 2) / 1024, 6);
    // A ceiling below every tried size never scales linearly: decline.
    const tiny = environment(2355, 10);
    expect(measureLocalFontNormalLine({ ...face } as FontFace, tiny.fonts, tiny.context, 2048, tiny.document)).toBeNull();
  });

  it('declines a strut whose computed family is not the measured face', () => {
    const face = { family: 'isolated-face', weight: '400', style: 'normal', status: 'loaded' } as FontFace;
    const env = environment();
    env.computed.family = 'Verdana';
    expect(measureLocalFontNormalLine(face, env.fonts, env.context, 2048, env.document)).toBeNull();
    expect(env.reads()).toBe(0);
  });

  it('declines unavailable, foreign-realm and invalid measurements', () => {
    const face = { family: 'isolated-face', weight: '400', style: 'normal', status: 'loaded' } as FontFace;
    const env = environment(Number.NaN);
    expect(measureLocalFontNormalLine(face, env.fonts, env.context, 2048, env.document)).toBeNull();
    expect(env.removed()).toBe(1);
    expect(measureLocalFontNormalLine(face, {} as FontFaceSet, env.context, 2048, env.document)).toBeNull();
    expect(measureLocalFontNormalLine(face, env.fonts, env.context, 2048, null)).toBeNull();
  });

  it('measures glyph ownership once per distinct cluster, not per distinct text', () => {
    const measured: string[] = [];
    vi.stubGlobal('OffscreenCanvas', class {
      getContext() {
        return {
          font: '',
          measureText(text: string) {
            measured.push(text);
            const owned = /^[a-z ]+$/.test(text);
            return { width: owned ? 0 : 50, actualBoundingBoxAscent: 0, actualBoundingBoxDescent: 0 } as TextMetrics;
          },
        };
      }
    });
    try {
      const metric = {
        family: 'isolated-face', requiresGlyphCoverage: true, glyphCoverageProbeFamily: 'isolated-face_coverage',
      } as ResolvedFontMetric;
      const route = { familyList: '"isolated-face", serif' } as CanvasFontRoute;
      expect(localFontMetricCoversText(metric, 'abc', route)).toBe(true);
      expect(localFontMetricCoversText(metric, 'cab ba', route)).toBe(true);
      expect(localFontMetricCoversText(metric, 'ab\u6f22', route)).toBe(false);
      // a, b, c, space and the CJK cluster: five native probes for three texts.
      expect(measured).toEqual(['a', 'b', 'c', ' ', '\u6f22']);
      // A joiner absorbing an unowned mark must not mark the joiner owned:
      // each cluster is probed on its own.
      measured.length = 0;
      expect(localFontMetricCoversText(metric, 'x\u2060\u0301', route)).toBe(false);
    } finally {
      vi.unstubAllGlobals();
    }
  });
});
