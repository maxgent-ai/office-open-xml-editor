import { describe, expect, it, vi } from 'vitest';
import {
  createFontResolver,
  type FontInventoryFace,
} from './font-service.js';
import {
  createTextLayoutService,
  TEXT_ROUTE_ORDINAL_LIMIT,
  textRouteOrdinalTableSize,
  type GlyphMeasureRequest,
  type GlyphMeasurement,
} from './text.js';

const faces: readonly FontInventoryFace[] = [
  { requestedFamily: 'Embedded Sans', resolvedFamily: 'Embedded Sans', source: 'embedded', weight: 400, style: 'normal' },
  { requestedFamily: 'Embedded Sans', resolvedFamily: 'Embedded Sans', source: 'embedded', weight: 700, style: 'italic' },
  { requestedFamily: 'Meiryo', resolvedFamily: '__ooxml_local_meiryo', source: 'local', weight: 400, style: 'normal' },
  { requestedFamily: 'Calibri', resolvedFamily: 'Carlito', source: 'google' },
  { requestedFamily: 'Legacy Arabic', resolvedFamily: 'Noto Naskh Arabic', source: 'substitute' },
  { requestedFamily: 'Theme Serif', resolvedFamily: 'Theme Serif', source: 'local' },
];

describe('font layout services', () => {
  it('recomputes an evicted shape with identical geometry and keeps service-scoped font invalidation', () => {
    const measure = vi.fn((request: Readonly<GlyphMeasureRequest>): GlyphMeasurement => ({
      advancePt: request.text.length * 2,
      ascentPt: 7,
      descentPt: 2,
    }));
    const service = createTextLayoutService({
      fonts: createFontResolver(faces),
      measurer: { fingerprint: 'cache-measure-v1', measure },
    });
    const shape = (text: string) => service.shape({ text, fontSizePt: 10, fonts: { ascii: 'Calibri' } });
    const first = shape('cached seed');
    const afterFirst = measure.mock.calls.length;
    expect(shape('cached seed')).toBe(first);
    expect(measure).toHaveBeenCalledTimes(afterFirst);

    for (let index = 0; index < 81922; index += 1) shape(`unique ${index}`);
    const beforeRevisit = measure.mock.calls.length;
    expect(shape('cached seed')).toEqual(first);
    expect(measure.mock.calls.length).toBeGreaterThan(beforeRevisit);

    const changed = createTextLayoutService({
      fonts: createFontResolver(faces),
      measurer: { fingerprint: 'cache-measure-v2', measure: (request) => ({
        advancePt: request.text.length * 3, ascentPt: 7, descentPt: 2,
      }) },
    });
    expect(changed.fingerprint).not.toBe(service.fingerprint);
    expect(changed.shape({ text: 'cached seed', fontSizePt: 10, fonts: { ascii: 'Calibri' } }).advancePt)
      .not.toBe(first.advancePt);
  }, 30_000);

  it('snapshots regional routes and includes their contents in the font fingerprint', () => {
    const routes = { sc: { Calibri: 'Carlito, "Noto Sans SC", sans-serif' } };
    const resolver = createFontResolver(faces, { regionalFamilyLists: routes });
    routes.sc.Calibri = 'Carlito, "Noto Sans TC", sans-serif';
    const changed = createFontResolver(faces, { regionalFamilyLists: routes });
    const request = { requestedFamily: 'Calibri', language: 'zh-CN' };
    expect(resolver.resolve(request).route.familyList).toContain('Noto Sans SC');
    expect(changed.resolve(request).route.familyList).toContain('Noto Sans TC');
    expect(resolver.fingerprint).not.toBe(changed.fingerprint);
  });

  it('shares one frozen resolution per exact request without aliasing distinct requests', () => {
    const routes = { jp: { Meiryo: '"Meiryo JP", sans-serif' }, sc: { Meiryo: '"Meiryo SC", sans-serif' } };
    const resolver = createFontResolver(faces, { regionalFamilyLists: routes });
    const first = resolver.resolve({ requestedFamily: 'Meiryo', weight: 700, language: 'ja-JP' });
    // Value-identical answers share their CSS family list and route instead of
    // retaining a fresh ~KB copy for every shaped span.
    expect(resolver.resolve({ requestedFamily: 'Meiryo', weight: 700, language: 'ja-JP' })).toBe(first);
    expect(Object.isFrozen(first) && Object.isFrozen(first.route)).toBe(true);
    expect(first.route.familyList).toContain('Meiryo JP');
    expect(resolver.resolve({ requestedFamily: 'Meiryo', weight: 700, language: 'zh-CN' }).route.familyList)
      .toContain('Meiryo SC');
    expect(resolver.resolve({ requestedFamily: 'Meiryo', weight: 400, language: 'ja-JP' }))
      .toMatchObject({ source: 'local', resolvedFamily: '__ooxml_local_meiryo' });
    expect(resolver.resolve({ requestedFamily: 'Meiryo', weight: 700, language: 'ja-JP', genericFamily: 'serif' }))
      .toMatchObject({ genericFamily: 'serif' });
  });

  it('keys glyph measurements by the exact route triple', () => {
    const measure = vi.fn((request: Readonly<GlyphMeasureRequest>): GlyphMeasurement => ({
      advancePt: request.text.length, ascentPt: 7, descentPt: 2,
    }));
    const routeFor = (familyList: string, scope: 'native' | 'generic', fingerprint: string) =>
      Object.freeze({ familyList, scope, fingerprint });
    let route = routeFor('"Same", sans-serif', 'native', 'route-a');
    const service = createTextLayoutService({
      fonts: {
        fingerprint: 'route-triple-v1',
        resolve: () => Object.freeze({
          requestedFamily: 'Same',
          resolvedFamily: 'Same',
          route,
          source: 'native' as const,
          weight: 400,
          style: 'normal' as const,
          diagnostics: [],
          genericFamily: 'sans-serif' as const,
        }),
      },
      measurer: { fingerprint: 'route-triple-measure-v1', measure },
    });
    // eastAsiaLanguage only selects a distinct shape entry here (Latin text),
    // so each call reaches the glyph-measurement cache.
    let variant = 0;
    const measuredOne = () => {
      service.shape({
        text: 'one', fontSizePt: 10, fonts: { ascii: 'Same' },
        clusterGeometry: false, eastAsiaLanguage: `x-${variant++}`,
      });
      return measure.mock.calls.filter(([request]) => request.text === 'one').length;
    };
    expect(measuredOne()).toBe(1);
    route = routeFor('"Same", sans-serif', 'native', 'route-a');
    expect(measuredOne()).toBe(1);
    route = routeFor('"Same", sans-serif', 'generic', 'route-a');
    expect(measuredOne()).toBe(2);
    route = routeFor('"Same", sans-serif', 'native', 'route-b');
    expect(measuredOne()).toBe(3);
    route = routeFor('"Other", sans-serif', 'native', 'route-a');
    expect(measuredOne()).toBe(4);
  });

  it('bounds route ordinals for ever-new routes without aliasing a retired route', () => {
    const measure = vi.fn((request: Readonly<GlyphMeasureRequest>): GlyphMeasurement => ({
      advancePt: request.fontRoute.familyList.length, ascentPt: 7, descentPt: 2,
    }));
    let family = 'Route 0';
    const service = createTextLayoutService({
      fonts: {
        fingerprint: 'unique-routes-v1',
        resolve: () => Object.freeze({
          requestedFamily: family,
          resolvedFamily: family,
          route: Object.freeze({
            familyList: `"${family}", sans-serif`,
            scope: 'native' as const,
            fingerprint: `route:${family}`,
          }),
          source: 'native' as const,
          weight: 400,
          style: 'normal' as const,
          diagnostics: [],
          genericFamily: 'sans-serif' as const,
        }),
      },
      measurer: { fingerprint: 'unique-routes-measure-v1', measure },
    });
    const advanceFor = (name: string) => {
      family = name;
      return service.shape({ text: 'x', fontSizePt: 10, fonts: { ascii: name }, clusterGeometry: false }).advancePt;
    };
    const first = advanceFor('Route 0');
    // Keep introducing new routes well past the table bound: the table stays
    // bounded instead of retaining every route ever measured.
    for (let index = 1; index < TEXT_ROUTE_ORDINAL_LIMIT * 2 + 10; index += 1) {
      advanceFor(`Route ${index}`);
      expect(textRouteOrdinalTableSize(service)).toBeLessThanOrEqual(TEXT_ROUTE_ORDINAL_LIMIT);
    }
    // After the reset, a route that received a re-issued ordinal and a retired
    // route are measured as themselves, never as each other.
    const reissued = advanceFor('Route with a much longer family name');
    expect(reissued).toBe('"Route with a much longer family name", sans-serif'.length);
    expect(advanceFor('Route 0')).toBe(first);
    expect(advanceFor('Route with a much longer family name')).toBe(reissued);
  });

  it('records embedded, local, Google, substitute, and generic resolution', () => {
    const resolver = createFontResolver(faces);

    expect(resolver.resolve({ requestedFamily: 'Embedded Sans', weight: 700, style: 'italic' }))
      .toMatchObject({ source: 'embedded', resolvedFamily: 'Embedded Sans', weight: 700, style: 'italic', diagnostics: [] });
    expect(resolver.resolve({ requestedFamily: 'Meiryo' }))
      .toMatchObject({ source: 'local', resolvedFamily: '__ooxml_local_meiryo' });
    expect(resolver.resolve({ requestedFamily: 'Meiryo', weight: 700 }))
      .toMatchObject({ source: 'native', resolvedFamily: 'Meiryo', genericFamily: 'sans-serif' });
    expect(resolver.resolve({ requestedFamily: 'Calibri' }))
      .toMatchObject({ source: 'google', resolvedFamily: 'Carlito' });

    const substituted = resolver.resolve({ requestedFamily: 'Legacy Arabic' });
    expect(substituted).toMatchObject({ source: 'substitute', resolvedFamily: 'Noto Naskh Arabic' });
    expect(substituted.diagnostics[0]?.message).toMatch(/implementation-dependent font substitution/i);

    const missing = resolver.resolve({ requestedFamily: 'Missing Face', genericFamily: 'serif' });
    expect(missing).toMatchObject({ source: 'native', resolvedFamily: 'Missing Face', genericFamily: 'serif' });
    expect(missing.diagnostics).toEqual([]);
  });

  it('selects font slots per script and shapes through the injected measurer', () => {
    const measured: GlyphMeasureRequest[] = [];
    const measure = (request: Readonly<GlyphMeasureRequest>): GlyphMeasurement => {
      measured.push({ ...request });
      return {
        advancePt: [...request.text].length * request.fontSizePt,
        ascentPt: request.fontSizePt * 0.8,
        descentPt: request.fontSizePt * 0.2,
      };
    };
    const service = createTextLayoutService({
      fonts: createFontResolver(faces),
      measurer: { fingerprint: 'fake-glyphs-v1', measure },
    });

    const result = service.shape({
      text: 'A国ش',
      fontSizePt: 10,
      weight: 700,
      style: 'italic',
      fonts: {
        ascii: 'Embedded Sans',
        highAnsi: 'Embedded Sans',
        eastAsia: 'Meiryo',
        complexScript: 'Legacy Arabic',
      },
    });

    expect(result.spans.map((span) => [span.text, span.font.source, span.font.resolvedFamily]))
      .toEqual([
        ['A', 'embedded', 'Embedded Sans'],
        ['国', 'native', 'Meiryo'],
        ['ش', 'embedded', 'Embedded Sans'],
      ]);
    expect(result.advancePt).toBe(30);
    expect(measured.map(({ fontRoute, weight, style }) => [fontRoute.familyList, weight, style]))
      .toEqual([
        ['"Embedded Sans", sans-serif', 700, 'italic'],
        ['"Meiryo", sans-serif', 700, 'italic'],
        ['"Embedded Sans", sans-serif', 700, 'italic'],
      ]);

    expect(createFontResolver(faces).resolve({
      requestedFamily: 'Embedded Sans',
      weight: 700,
      style: 'normal',
    })).toMatchObject({ source: 'native', resolvedFamily: 'Embedded Sans' });

    const forced = service.shape({
      text: 'ش-12',
      fontSizePt: 10,
      complexScript: true,
      fonts: { ascii: 'Embedded Sans', complexScript: 'Legacy Arabic' },
    });
    expect(forced.spans).toHaveLength(1);
    expect(forced.spans[0]).toMatchObject({ text: 'ش-12', script: 'complexScript' });
    expect(forced.spans[0]?.font).toMatchObject({ source: 'substitute', resolvedFamily: 'Noto Naskh Arabic' });
  });

  it('retains aggregate glyph ink bounds independently from typographic metrics', () => {
    const service = createTextLayoutService({
      fonts: createFontResolver(faces),
      measurer: {
        fingerprint: 'ink-bounds-v1',
        measure: (request) => request.text === 'A'
          ? {
              advancePt: 8, ascentPt: 9, descentPt: 3,
              inkBounds: { xMinPt: -1, xMaxPt: 7, ascentPt: 7, descentPt: 2 },
            }
          : {
              advancePt: 6, ascentPt: 10, descentPt: 4,
              inkBounds: { xMinPt: 0, xMaxPt: 8, ascentPt: 8, descentPt: 1 },
            },
      },
    });

    const result = service.shape({
      text: 'A国', fontSizePt: 10,
      fonts: { ascii: 'Embedded Sans', eastAsia: 'Meiryo' },
    });

    expect(result.inkBounds).toEqual({
      xMinPt: -1,
      xMaxPt: 16,
      ascentPt: 8,
      descentPt: 2,
    });
    expect(result.spans.map((span) => span.inkBounds)).toEqual([
      { xMinPt: -1, xMaxPt: 7, ascentPt: 7, descentPt: 2 },
      { xMinPt: 0, xMaxPt: 8, ascentPt: 8, descentPt: 1 },
    ]);
    expect(structuredClone(result.inkBounds)).toEqual(result.inkBounds);
  });

  it('uses theme slots without collapsing mixed-script runs to one family', () => {
    const service = createTextLayoutService({
      fonts: createFontResolver(faces),
      measurer: {
        fingerprint: 'fake-glyphs-v1',
        measure: (request) => ({
          advancePt: request.text.length,
          ascentPt: 1,
          descentPt: 0,
        }),
      },
    });

    const result = service.shape({
      text: 'A国',
      fontSizePt: 12,
      fonts: {},
      themeFonts: { ascii: 'Theme Serif', eastAsia: 'Meiryo' },
    });

    expect(result.spans.map((span) => span.font.requestedFamily))
      .toEqual(['Theme Serif', 'Meiryo']);
    expect(service.fingerprint).toMatch(/^text:/);
  });

  it('gives each theme axis precedence over its matching direct axis', () => {
    const service = createTextLayoutService({
      fonts: createFontResolver(faces),
      measurer: {
        fingerprint: 'theme-precedence-v1',
        measure: (request) => ({ advancePt: request.text.length, ascentPt: 1, descentPt: 0 }),
      },
    });

    const result = service.shape({
      text: 'Aé国ش',
      fontSizePt: 12,
      complexScript: true,
      fonts: {
        ascii: 'Embedded Sans',
        highAnsi: 'Embedded Sans',
        eastAsia: 'Embedded Sans',
        complexScript: 'Embedded Sans',
      },
      themeFonts: {
        ascii: 'Theme Serif',
        highAnsi: 'Calibri',
        eastAsia: 'Meiryo',
        complexScript: 'Legacy Arabic',
      },
    });

    expect(result.spans.map((span) => [span.script, span.font.requestedFamily]))
      .toEqual([
        ['complexScript', 'Legacy Arabic'],
      ]);

    const perAxis = ['A', 'é', '国'].map((text) => service.shape({
      text,
      fontSizePt: 12,
      fonts: {
        ascii: 'Embedded Sans',
        highAnsi: 'Embedded Sans',
        eastAsia: 'Embedded Sans',
      },
      themeFonts: {
        ascii: 'Theme Serif',
        highAnsi: 'Calibri',
        eastAsia: 'Meiryo',
      },
    }).spans[0]?.font.requestedFamily);
    expect(perAxis).toEqual(['Theme Serif', 'Calibri', 'Meiryo']);
  });

  it('keeps an authored but unresolved theme attribute suppressing the direct face on every axis', () => {
    const service = createTextLayoutService({
      fonts: createFontResolver([]),
      measurer: {
        fingerprint: 'unresolved-theme-v1',
        measure: (request) => ({ advancePt: request.text.length, ascentPt: 1, descentPt: 0 }),
      },
    });
    const fonts = {
      ascii: 'Direct ASCII',
      highAnsi: 'Direct HANSI',
      eastAsia: 'Direct EA',
      complexScript: 'Direct CS',
    };
    const themeFontPresence = {
      ascii: true,
      highAnsi: true,
      eastAsia: true,
      complexScript: true,
    };

    for (const slot of ['ascii', 'highAnsi', 'eastAsia', 'complexScript'] as const) {
      const resolved = service.resolve({
        fonts,
        themeFonts: {},
        themeFontPresence,
        slot,
      } as Parameters<typeof service.resolve>[0]);
      const generic = slot === 'eastAsia' ? 'sans-serif' : 'serif';
      expect(resolved, slot).toMatchObject({
        source: 'generic', requestedFamily: generic, genericFamily: generic,
      });
    }
  });

  it('snapshots charset routing so caller mutation cannot change shape under one fingerprint', () => {
    const charsets = { 'ea face': '86' };
    const service = createTextLayoutService({
      fonts: createFontResolver([]),
      eastAsiaFontCharsets: charsets,
      measurer: {
        fingerprint: 'charset-snapshot-v1',
        measure: (request) => ({ advancePt: request.text.length, ascentPt: 1, descentPt: 0 }),
      },
    });
    const shape = () => service.shape({
      text: '\u0100',
      fontSizePt: 10,
      fontHint: 'eastAsia',
      fonts: { highAnsi: 'HANSI Face', eastAsia: 'EA Face' },
    }).spans[0]?.script;
    const fingerprint = service.fingerprint;

    expect(shape()).toBe('eastAsia');
    charsets['ea face'] = '00';
    expect(shape()).toBe('eastAsia');
    expect(service.fingerprint).toBe(fingerprint);
  });

  it('canonicalizes exact face tuples independently of inventory order', () => {
    const inventory: FontInventoryFace[] = [
      { requestedFamily: 'Tuple Face', resolvedFamily: 'Tuple Face', source: 'embedded', weight: 700, style: 'italic' },
      { requestedFamily: 'Tuple Face', resolvedFamily: 'Tuple Face', source: 'embedded', weight: 400, style: 'normal' },
      { requestedFamily: 'Tuple Face', resolvedFamily: 'Tuple Face', source: 'embedded', weight: 700, style: 'normal' },
    ];

    expect(createFontResolver(inventory).fingerprint)
      .toBe(createFontResolver([...inventory].reverse()).fingerprint);
  });

  it('produces identical main and worker fingerprints from identical snapshots', () => {
    const make = () => createTextLayoutService({
      fonts: createFontResolver([...faces].reverse()),
      measurer: {
        fingerprint: 'canvas-text-metrics-v1',
        measure: () => ({ advancePt: 1, ascentPt: 1, descentPt: 0 }),
      },
    });

    const main = make();
    const worker = make();
    expect(main.fingerprint).toBe(worker.fingerprint);
    expect(main.shape({ text: 'A国', fontSizePt: 10, fonts: { ascii: 'Calibri', eastAsia: 'Meiryo' } }))
      .toEqual(worker.shape({ text: 'A国', fontSizePt: 10, fonts: { ascii: 'Calibri', eastAsia: 'Meiryo' } }));
  });

  it('classifies every Unicode scalar without splitting surrogate pairs', () => {
    const service = createTextLayoutService({
      fonts: createFontResolver(faces),
      measurer: {
        fingerprint: 'cluster-measurer-v1',
        measure: (request) => ({ advancePt: [...request.text].length, ascentPt: 1, descentPt: 0 }),
      },
    });
    const shape = (text: string) => service.shape({
      text,
      fontSizePt: 10,
      fonts: {
        ascii: 'Embedded Sans',
        highAnsi: 'Theme Serif',
        eastAsia: 'Meiryo',
        complexScript: 'Legacy Arabic',
      },
    });

    expect(shape('a\u0301').spans.map((span) => [span.text, span.script]))
      .toEqual([['a', 'ascii'], ['\u0301', 'highAnsi']]);
    expect(shape('a\u0301').graphemeBoundaries).toEqual([0, 2]);
    expect(shape('a\u0301').clusters).toEqual([
      { range: { start: 0, end: 2 }, offsetPt: 0, advancePt: 2 },
    ]);
    expect(shape('a\u0301').spans.map((span) => span.breakBefore)).toEqual([true, false]);
    expect(shape('国\u{E0100}').spans.map((span) => [span.text, span.script]))
      .toEqual([['国\u{E0100}', 'eastAsia']]);
    expect(shape('\u{20000}').spans.map((span) => [span.text, span.script]))
      .toEqual([['\u{20000}', 'eastAsia']]);
    expect(shape('\u{20000}観').clusters).toEqual([
      { range: { start: 0, end: 2 }, offsetPt: 0, advancePt: 1 },
      { range: { start: 2, end: 3 }, offsetPt: 1, advancePt: 1 },
    ]);
    expect(shape('👩‍💻').spans.map((span) => [span.text, span.script]))
      .toEqual([['👩', 'eastAsia'], ['\u200d', 'highAnsi'], ['💻', 'eastAsia']]);
    expect(shape('𠀀').spans.flatMap((span) => [...span.text])).toEqual(['𠀀']);
    expect(shape('ش-12').spans.every((span) => span.script !== 'complexScript')).toBe(true);
  });

  it('measures each contextual grapheme boundary once', () => {
    const measured: string[] = [];
    const service = createTextLayoutService({
      fonts: createFontResolver(faces),
      measurer: {
        fingerprint: 'contextual-boundary-count-v1',
        measure: (request) => {
          measured.push(request.text);
          return {
            advancePt: request.text.length ** 2,
            ascentPt: 1,
            descentPt: 0,
          };
        },
      },
    });

    const request = {
      text: 'abcd',
      fontSizePt: 10,
      fonts: { ascii: 'Embedded Sans' },
    } as const;
    const shaped = service.shape(request);

    // Cluster advances remain contextual prefix differences. The acquisition
    // boundary must not ask the native measurer for the same prefix twice just
    // because it is both one cluster's end and the next cluster's start.
    expect(shaped.clusters).toEqual([
      { range: { start: 0, end: 1 }, offsetPt: 0, advancePt: 1 },
      { range: { start: 1, end: 2 }, offsetPt: 1, advancePt: 3 },
      { range: { start: 2, end: 3 }, offsetPt: 4, advancePt: 5 },
      { range: { start: 3, end: 4 }, offsetPt: 9, advancePt: 7 },
    ]);
    expect(measured).toEqual(['abcd', 'a', 'ab', 'abc']);
  });

  it('does not retain contextual grapheme prefixes in the document measurement cache', () => {
    const measured: string[] = [];
    const service = createTextLayoutService({
      fonts: createFontResolver(faces),
      measurer: {
        fingerprint: 'contextual-prefix-cache-scope-v1',
        measure: (request) => {
          measured.push(request.text);
          return {
            advancePt: request.text.length,
            ascentPt: 1,
            descentPt: 0,
          };
        },
      },
    });
    const text = 'a'.repeat(1_024);
    service.shape({ text, fontSizePt: 10, fonts: { ascii: 'Embedded Sans' } });
    expect(measured).toContain(text.slice(0, 512));

    // A single script run may retain its complete request, but contextual
    // prefixes are temporary facts owned by that shape call's prefixAdvances
    // map: a later whole-span measurement of an earlier prefix must miss the
    // document measurement cache.
    measured.length = 0;
    service.shape({
      text: text.slice(0, 512),
      fontSizePt: 10,
      fonts: { ascii: 'Embedded Sans' },
      clusterGeometry: false,
    });
    expect(measured).toEqual([text.slice(0, 512)]);
  });

  it('reuses identical glyph measurements across convergence shape calls', () => {
    const measured: string[] = [];
    const service = createTextLayoutService({
      fonts: createFontResolver(faces),
      measurer: {
        fingerprint: 'convergence-measure-cache-v1',
        measure: (request) => {
          measured.push(request.text);
          return {
            advancePt: request.text.length,
            ascentPt: 1,
            descentPt: 0,
          };
        },
      },
    });
    const request = {
      text: 'repeat',
      fontSizePt: 10,
      fonts: { ascii: 'Embedded Sans' },
    } as const;

    expect(service.shape(request)).toBe(service.shape(request));

    expect(measured).toEqual(['repeat', 'r', 're', 'rep', 'repe', 'repea']);
  });

  it('reuses a complete glyph request across distinct shape acquisitions', () => {
    const measured: string[] = [];
    const service = createTextLayoutService({
      fonts: createFontResolver(faces),
      measurer: {
        fingerprint: 'complete-measure-cache-v1',
        measure: (request) => {
          measured.push(request.text);
          return {
            advancePt: request.text.length,
            ascentPt: 1,
            descentPt: 0,
          };
        },
      },
    });
    const request = {
      text: 'repeat',
      fontSizePt: 10,
      fonts: { ascii: 'Embedded Sans' },
      clusterGeometry: false,
    } as const;

    const first = service.shape(request);
    const second = service.shape({ ...request, eastAsiaLanguage: 'en-US' });

    expect(first).not.toBe(second);
    expect(first).toEqual(second);
    expect(measured).toEqual(['repeat']);
  });

  it('can acquire aggregate metrics without contextual cluster geometry', () => {
    const measured: string[] = [];
    const service = createTextLayoutService({
      fonts: createFontResolver(faces),
      measurer: {
        fingerprint: 'aggregate-only-v1',
        measure: (request) => {
          measured.push(request.text);
          return { advancePt: request.text.length, ascentPt: 1, descentPt: 0 };
        },
      },
    });

    const shaped = service.shape({
      text: 'abcd',
      fontSizePt: 10,
      fonts: { ascii: 'Embedded Sans' },
      clusterGeometry: false,
    });

    expect(shaped.advancePt).toBe(4);
    expect(shaped.clusters).toBeUndefined();
    expect(measured).toEqual(['abcd']);
  });

  it('implements the complete ECMA-376 §17.3.2.26 character-range table', () => {
    const service = createTextLayoutService({
      fonts: createFontResolver(faces),
      measurer: {
        fingerprint: 'script-slot-table-v1',
        measure: (request) => ({ advancePt: [...request.text].length, ascentPt: 1, descentPt: 0 }),
      },
    });
    const slot = (text: string, complexScript = false) => service.shape({
      text,
      fontSizePt: 10,
      complexScript,
      fonts: {
        ascii: 'Embedded Sans',
        highAnsi: 'Theme Serif',
        eastAsia: 'Meiryo',
        complexScript: 'Legacy Arabic',
      },
    }).spans.map((span) => span.script);

    const cases: Array<[string, string, 'ascii' | 'highAnsi' | 'eastAsia']> = [
      ['ASCII', 'A', 'ascii'],
      ['Latin-1', 'é', 'highAnsi'],
      ['Hebrew', '\u05D0', 'ascii'],
      ['Arabic', '\u0627', 'ascii'],
      ['Syriac', '\u0710', 'ascii'],
      ['Thaana', '\u0780', 'ascii'],
      ['Hebrew presentation form', '\uFB1D', 'ascii'],
      ['Arabic presentation form A', '\uFB50', 'ascii'],
      ['Arabic presentation form B', '\uFE70', 'ascii'],
      ['Hangul Jamo', '\u1100', 'eastAsia'],
      ['Yi syllable', '\uA000', 'eastAsia'],
      ['CJK compatibility form', '\uFE30', 'eastAsia'],
      ['small form variant', '\uFE50', 'eastAsia'],
    ];
    for (const [name, text, expected] of cases) {
      expect(slot(text), name).toEqual([expected]);
    }
    expect(slot('\u05D0', true)).toEqual(['complexScript']);
    expect(slot('\u0627', true)).toEqual(['complexScript']);
  });

  it('implements the conditional eastAsia hint, language, charset, and cs precedence rows', () => {
    const service = createTextLayoutService({
      fonts: createFontResolver(faces),
      measurer: {
        fingerprint: 'conditional-script-slot-table-v1',
        measure: (request) => ({ advancePt: [...request.text].length, ascentPt: 1, descentPt: 0 }),
      },
    });
    const slot = (
      text: string,
      options: {
        hint?: 'default' | 'eastAsia' | 'cs';
        eastAsiaLanguage?: string;
        eastAsiaFontCharset?: string;
        complexScript?: boolean;
      } = {},
    ) => service.shape({
      text,
      fontSizePt: 10,
      fontHint: options.hint,
      eastAsiaLanguage: options.eastAsiaLanguage,
      eastAsiaFontCharset: options.eastAsiaFontCharset,
      complexScript: options.complexScript,
      fonts: {
        ascii: 'Embedded Sans',
        highAnsi: 'Theme Serif',
        eastAsia: 'Meiryo',
        complexScript: 'Legacy Arabic',
      },
    }).spans[0]?.script;

    const eastAsia = { hint: 'eastAsia' as const };
    expect(slot('\u00a1', eastAsia)).toBe('eastAsia');
    expect(slot('\u00e0', { ...eastAsia, eastAsiaLanguage: 'zh-CN' })).toBe('eastAsia');
    expect(slot('\u00e0', { ...eastAsia, eastAsiaLanguage: 'ja-JP' })).toBe('highAnsi');
    expect(slot('\u0100', { ...eastAsia, eastAsiaLanguage: 'zh-TW' })).toBe('eastAsia');
    expect(slot('\u0180', { ...eastAsia, eastAsiaFontCharset: '88' })).toBe('eastAsia');
    expect(slot('\u0250', { ...eastAsia, eastAsiaFontCharset: '86' })).toBe('eastAsia');
    expect(slot('\u0100', { ...eastAsia, eastAsiaFontCharset: '80' })).toBe('highAnsi');
    for (const scalar of ['\u02b0', '\u0300', '\u0370', '\u0400', '\u2000', '\u20a0', '\u2100', '\u2190', '\u2200', '\u2300', '\u2400', '\u2500', '\u2600', '\u2700', '\ue000', '\ufb00']) {
      expect(slot(scalar, eastAsia), scalar).toBe('eastAsia');
      expect(slot(scalar), scalar).toBe('highAnsi');
    }
    expect(slot('\u1e00', { ...eastAsia, eastAsiaLanguage: 'zh-Hans' })).toBe('eastAsia');
    expect(slot('\u1e00', { ...eastAsia, eastAsiaLanguage: 'ko-KR' })).toBe('highAnsi');

    // ECMA-376 Part 1 (5th ed.) §17.3.2.26: listed endpoints are inclusive
    // and every unlisted gap falls back to hAnsi.
    const boundaries: Array<[
      string,
      'ascii' | 'highAnsi' | 'eastAsia',
      { hint: 'eastAsia' }?,
    ]> = [
      ['\u02ff', 'eastAsia', eastAsia],
      ['\u0300', 'eastAsia', eastAsia],
      ['\u036f', 'eastAsia', eastAsia],
      ['\u0370', 'eastAsia', eastAsia],
      ['\u03cf', 'eastAsia', eastAsia],
      ['\u03d0', 'highAnsi', eastAsia],
      ['\u03ff', 'highAnsi', eastAsia],
      ['\u0400', 'eastAsia', eastAsia],
      ['\u2e80', 'eastAsia'],
      ['\u2e80', 'eastAsia', eastAsia],
      ['\u2eff', 'eastAsia', eastAsia],
      ['\u2f00', 'eastAsia'],
      ['\u2fdf', 'eastAsia'],
      ['\u2fe0', 'highAnsi'],
      ['\u2fef', 'highAnsi'],
      ['\u2ff0', 'eastAsia'],
      ['\u318f', 'eastAsia'],
      ['\u3190', 'eastAsia'],
      ['\u319f', 'eastAsia'],
      ['\u31a0', 'highAnsi'],
      ['\u31ff', 'highAnsi'],
      ['\u3200', 'eastAsia'],
      ['\u4dbf', 'eastAsia'],
      ['\u4dc0', 'highAnsi'],
      ['\u4dff', 'highAnsi'],
      ['\u4e00', 'eastAsia'],
      ['\u9faf', 'eastAsia'],
      ['\u9fb0', 'highAnsi'],
      ['\ua4cf', 'eastAsia'],
      ['\ua4d0', 'highAnsi'],
      ['\ud7af', 'eastAsia'],
      ['\ud7b0', 'highAnsi'],
      ['\ufafe', 'eastAsia'],
      ['\ufaff', 'eastAsia'],
      ['\ufe6f', 'eastAsia'],
      ['\ufe70', 'ascii'],
      ['\ufefe', 'ascii'],
      ['\ufeff', 'highAnsi'],
      ['\u{1f642}', 'eastAsia'],
    ];
    for (const [scalar, expected, options] of boundaries) {
      expect(slot(scalar, options), `U+${scalar.codePointAt(0)?.toString(16)}`).toBe(expected);
    }

    // Step 2: cs/rtl wins unless step 1 selected eastAsia while hint=eastAsia.
    expect(slot('\u4e00', { complexScript: true })).toBe('complexScript');
    expect(slot('\u4e00', { ...eastAsia, complexScript: true })).toBe('eastAsia');
    expect(slot('\u2e80', { complexScript: true })).toBe('complexScript');
    expect(slot('\u2e80', { ...eastAsia, complexScript: true })).toBe('eastAsia');
    expect(slot('A', { ...eastAsia, complexScript: true })).toBe('complexScript');
  });

  it('includes immutable local geometry metrics in the text fingerprint', () => {
    const make = (lineHeightRatio: number) => createTextLayoutService({
      fonts: createFontResolver(faces),
      localMetrics: {
        authored: { family: '__ooxml_local_authored', lineHeightRatio },
      },
      measurer: {
        fingerprint: 'metrics-v1',
        measure: () => ({ advancePt: 1, ascentPt: 1, descentPt: 0 }),
      },
    });

    expect(make(1.3).fingerprint).not.toBe(make(1.31).fingerprint);
    expect(Object.isFrozen(make(1.3).localMetrics.authored)).toBe(true);
  });
});
