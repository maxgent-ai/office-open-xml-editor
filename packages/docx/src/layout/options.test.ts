import { describe, expect, it } from 'vitest';
import { createLayoutOptionsKeyer, normalizeLayoutOptions } from './options.js';
import type { LayoutServices } from './types.js';
import { createCanvasFontRoute } from '@silurus/ooxml-core';

function services(
  text: string,
  images: string,
  math: string,
  verticalGlyphs = 'vertical:a',
): LayoutServices {
  return {
    text: {
      fingerprint: text,
      localMetrics: {},
      resolve: () => ({
        requestedFamily: 'sans-serif', resolvedFamily: 'sans-serif',
        route: createCanvasFontRoute('sans-serif', 'generic'),
        source: 'generic', weight: 400, style: 'normal', diagnostics: [], genericFamily: 'sans-serif',
      }),
      shape: () => ({ advancePt: 0, ascentPt: 0, descentPt: 0, spans: [], graphemeBoundaries: [0], diagnostics: [] }),
    },
    images: { fingerprint: images, resolve: () => ({ widthPt: 1, heightPt: 1, mimeType: 'image/png' }) },
    math: { fingerprint: math, resolve: () => ({ resourceKey: 'm', widthEm: 1, ascentEm: 1, descentEm: 0, diagnostics: [] }) },
    verticalGlyphFingerprint: verticalGlyphs,
  };
}

describe('layout options', () => {
  it('normalizes Date, number, and undefined against one captured load-time default', () => {
    expect(normalizeLayoutOptions(new Date(123), 999)).toEqual({ currentDateMs: 123 });
    expect(normalizeLayoutOptions(456, 999)).toEqual({ currentDateMs: 456 });
    expect(normalizeLayoutOptions(undefined, 999)).toEqual({ currentDateMs: 999 });

    if (false) {
      // @ts-expect-error environment strings are not a layout input
      normalizeLayoutOptions(undefined, 'browser-fonts-v1');
    }
  });

  it('keys only the normalized date and actual service fingerprints', () => {
    const layoutOptionsKey = createLayoutOptionsKeyer();
    const base = services('text:a', 'images:a', 'math:a');
    const key = layoutOptionsKey({ currentDateMs: 100 }, base);

    expect(layoutOptionsKey({ currentDateMs: 101 }, base)).not.toBe(key);
    expect(layoutOptionsKey({ currentDateMs: 100, showTrackedChanges: true }, base)).not.toBe(key);
    expect(layoutOptionsKey({ currentDateMs: 100 }, services('text:b', 'images:a', 'math:a'))).not.toBe(key);
    expect(layoutOptionsKey({ currentDateMs: 100 }, services('text:a', 'images:b', 'math:a'))).not.toBe(key);
    expect(layoutOptionsKey({ currentDateMs: 100 }, services('text:a', 'images:a', 'math:b'))).not.toBe(key);
    expect(layoutOptionsKey({ currentDateMs: 100 }, services(
      'text:a', 'images:a', 'math:a', 'vertical:b',
    ))).not.toBe(key);
    // A fingerprint moved to another service is a different input.
    expect(layoutOptionsKey({ currentDateMs: 100 }, services('images:a', 'text:a', 'math:a'))).not.toBe(key);

    if (false) {
      // @ts-expect-error paint width, DPR, and color cannot enter the layout key
      layoutOptionsKey({ currentDateMs: 100 }, base, { width: 600, dpr: 2, defaultTextColor: '#fff' });
    }
  });

  it('interns service fingerprints so a key never spells them out', () => {
    const layoutOptionsKey = createLayoutOptionsKeyer();
    // A realistic text-service fingerprint embeds the document's font-metric
    // snapshot; equal content in a distinct string must select the same key.
    const metrics = 'x'.repeat(55_000);
    const key = layoutOptionsKey({ currentDateMs: 100 }, services(`text:${metrics}`, 'images:a', 'math:a'));
    const equal = layoutOptionsKey(
      { currentDateMs: 100 },
      services(['text:', metrics].join(''), 'images:a', 'math:a'),
    );

    expect(equal).toBe(key);
    expect(key.length).toBeLessThan(200);
    expect(layoutOptionsKey({ currentDateMs: 100 }, services(`text:${metrics}y`, 'images:a', 'math:a')))
      .not.toBe(key);
  });
});
