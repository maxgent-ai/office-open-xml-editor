import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { chromium, firefox, webkit, type Browser, type Page } from '@playwright/test';
import { createServer, type ViteDevServer } from 'vite';
import { existsSync } from 'node:fs';
import { storeZip } from './conformance/generate.js';

type Api = typeof import('./native-route-line-metrics') & typeof import('./layout-runtime')
  & typeof import('./line-layout') & typeof import('./index') & typeof import('@silurus/ooxml-core');

// An authored family that no host installs. Its native list is the face's
// name followed by the sans fallback tail; the face that paints is whichever
// tail family the host has first. The assertions below are conditional on
// that face being Arial, whose catalogued hhea line the CSS normal line of
// Chromium on macOS reproduces; other hosts skip the numeric checks.
const ABSENT = 'Ooxml Absent Sans 1843';

describe.skipIf(!existsSync(chromium.executablePath()))('DOCX native route line geometry', () => {
  let server: ViteDevServer;
  let browser: Browser;
  let page: Page;

  beforeAll(async () => {
    server = await createServer({
      configFile: new URL('../vite.config.ts', import.meta.url).pathname,
      server: { host: '127.0.0.1', port: 6685, strictPort: true },
      logLevel: 'error',
    });
    await server.listen();
    browser = await chromium.launch();
    page = await browser.newPage();
    await page.addInitScript(() => { (window as unknown as { __OOXML_MODEL_SOURCES__: boolean }).__OOXML_MODEL_SOURCES__ = true; });
    await page.route('**/native-route.html', route => route.fulfill({
      contentType: 'text/html', body: '<!doctype html><body></body>',
    }));
    await page.goto('http://127.0.0.1:6685/native-route.html');
    await page.addScriptTag({ type: 'module', content: `
      import * as native from '/src/native-route-line-metrics.ts';
      import * as runtime from '/src/layout-runtime.ts';
      import * as lines from '/src/line-layout.ts';
      import { DocxDocument } from '/src/index.ts';
      import * as core from ${JSON.stringify('/@fs' + new URL('../../core/src/index.ts', import.meta.url).pathname)};
      window.nativeRouteTest = { ...core, ...native, ...runtime, ...lines, DocxDocument };
    ` });
    await page.waitForFunction(() => 'nativeRouteTest' in window);
  }, 60_000);
  afterAll(async () => { await browser?.close(); await server?.close(); });

  it('admits the painting face line only for clusters that face draws', async (test) => {
    const result = await page.evaluate(async (absent) => {
      const api = (window as unknown as { nativeRouteTest: Api }).nativeRouteTest;
      const model = {
        section: { pageWidth: 612, pageHeight: 792, marginTop: 72, marginRight: 72,
          marginBottom: 72, marginLeft: 72, headerDistance: 36, footerDistance: 36 },
        body: [], headers: {}, footers: {}, fontFamilyClasses: { [absent]: 'swiss' },
      } as unknown as import('./types').DocxDocumentModel;
      const ctx = document.createElement('canvas').getContext('2d') as CanvasRenderingContext2D;
      const plain = api.createLayoutServices(model, { measureContext: ctx });
      const route = plain.text.resolve({ fonts: { ascii: absent }, slot: 'ascii', weight: 400, style: 'normal' });
      const prepared = await api.measureDocxNativeRouteLines([{
        familyList: route.route.familyList, weight: 400, style: 'normal',
      }], true);
      const line = prepared.lines[0];
      const admitted = api.admitDocxNativeRouteLines(prepared.lines, document.fonts);
      const measured = api.createLayoutServices(model, { measureContext: ctx, nativeRouteMetrics: admitted });
      const values = ['ABC', '漢字', 'ABC漢字', 'e\u0301', '\u0905\u0915', '\u2764\ufe0f'].map(text => {
        const run = { type: 'text' as const, text, fontFamily: absent, fontFamilyHighAnsi: absent,
          fontFamilyEastAsia: absent, fontFamilyCs: absent, fontSize: 20, bold: false, italic: false,
          underline: false, strikethrough: false, color: null, isLink: false,
          background: null, vertAlign: null, hyperlink: null };
        return [measured, plain].map(layoutServices => {
          const segments = api.buildSegments([run], { pageIndex: 0, totalPages: 1, layoutServices });
          const lines = api.layoutLines(ctx, segments, 400, 0, 1);
          return {
            owned: segments.filter(segment => 'text' in segment).map(segment => ({
              text: segment.text, resource: segment.resolvedResourceVerticalMetric === true,
            })),
            sides: lines.map(entry => [entry.ascent, entry.descent]),
          };
        });
      });
      api.unloadOfficeFontFallbacks(prepared.faces);
      return { primary: line?.primaryFamily ?? null, admitted: Object.keys(admitted).length, values };
    }, ABSENT);
    if (result.primary !== 'Arial' || result.admitted !== 1) test.skip();
    const [latin, han, mixed, combining, devanagari, emoji] = result.values;
    expect(latin![0].owned.every(value => value.resource)).toBe(true);
    // 20 pt Arial: (1854 + 434 + 67) / 2048 em, the face's own normal line.
    const extent = (sides: number[][]) => sides[0]![0]! + sides[0]![1]!;
    expect(extent(latin![0].sides)).toBeCloseTo(20 * 2355 / 2048, 6);
    expect(extent(latin![1].sides)).toBeLessThan(extent(latin![0].sides));
    // Han and Devanagari fall through to later faces: no borrowed geometry,
    // exactly the Canvas-box result of the route without a measurement.
    for (const value of [han!, devanagari!, emoji!]) {
      expect(value[0].owned.every(entry => !entry.resource)).toBe(true);
      expect(value[0].sides).toEqual(value[1].sides);
    }
    expect(mixed![0].owned.some(value => value.text.includes('ABC') && value.resource)).toBe(true);
    expect(mixed![0].owned.filter(value => !value.text.includes('ABC')).every(value => !value.resource)).toBe(true);
    expect(combining![0].owned.every(value => value.resource)).toBe(true);
  });

  it('paginates identically in main and worker mode with the measured line', async (test) => {
    const encoder = new TextEncoder();
    const parts = new Map<string, Uint8Array>();
    parts.set('[Content_Types].xml', encoder.encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'));
    parts.set('_rels/.rels', encoder.encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'));
    const run = (text: string, bold = false) => `<w:r><w:rPr><w:rFonts w:ascii="${ABSENT}" w:hAnsi="${ABSENT}" w:eastAsia="${ABSENT}"/>${bold ? '<w:b/>' : ''}<w:sz w:val="22"/></w:rPr><w:t>${text}</w:t></w:r>`;
    const paragraph = (body: string) => `<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>${body}</w:p>`;
    const body = [
      paragraph(run('Absent face pagination')).repeat(120),
      paragraph(run('Bold absent face', true)).repeat(20),
      paragraph(run('漢字')).repeat(20),
    ].join('');
    parts.set('word/document.xml', encoder.encode(`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`));
    const bytes = Array.from(storeZip(parts));
    const result = await page.evaluate(async ({ bytes, absent }) => {
      const api = (window as unknown as { nativeRouteTest: Api }).nativeRouteTest;
      const values = [];
      for (const mode of ['main', 'worker'] as const) {
        const doc = await api.DocxDocument.load(new Uint8Array(bytes).buffer, { mode });
        try {
          const runs = [];
          for (let index = 0; index < doc.pageCount; index++) runs.push(await doc.collectPageRuns(index));
          values.push({ pages: doc.pageCount, runs });
        } finally { doc.destroy(); }
      }
      // Independent identification of the painting face for the expectation.
      const context = document.createElement('canvas').getContext('2d') as CanvasRenderingContext2D;
      context.font = `100px "${absent}", "Noto Sans", "Noto Sans Hebrew", "Noto Sans Thai", "Arial"`;
      const listBox = context.measureText('');
      context.font = '100px "Arial"';
      const arialBox = context.measureText('');
      const arialPaints = listBox.fontBoundingBoxAscent === arialBox.fontBoundingBoxAscent
        && listBox.fontBoundingBoxDescent === arialBox.fontBoundingBoxDescent;
      return { values, arialPaints };
    }, { bytes, absent: ABSENT });
    expect(result.values[1]).toEqual(result.values[0]);
    if (!result.arialPaints) test.skip();
    expect(result.values[0].runs[0][0].h).toBeCloseTo(11 * 4 / 3 * 2355 / 2048, 6);
  }, 60_000);

  // The oracle itself, in every engine: clusters an engine draws outside the
  // family list (emoji presentation, unsupported variation sequences) or with
  // a zero-advance fallback mark must never count as owned.
  for (const engine of [chromium, firefox, webkit]) {
    it(`judges cluster ownership soundly in ${engine.name()}`, async (test) => {
      if (!existsSync(engine.executablePath())) test.skip();
      const other = await engine.launch();
      try {
        const enginePage = await other.newPage();
        await enginePage.route('**/native-route.html', route => route.fulfill({
          contentType: 'text/html', body: '<!doctype html><body></body>',
        }));
        await enginePage.goto('http://127.0.0.1:6685/native-route.html');
        await enginePage.addScriptTag({ type: 'module', content: `
          import * as oracle from ${JSON.stringify('/@fs' + new URL('../../core/src/fonts/native-route-line-metrics.ts', import.meta.url).pathname)};
          window.oracle = oracle;
        ` });
        await enginePage.waitForFunction(() => 'oracle' in window, undefined, { timeout: 30_000 });
        const result = await enginePage.evaluate(async (absent) => {
          const oracle = (window as unknown as { oracle: Pick<typeof import('@silurus/ooxml-core'),
            'loadNativeRouteSentinels' | 'measureNativeRouteNormalLines'> & {
            nativeRouteOwnershipFonts(list: string, weight: number, style: 'normal'): readonly [string, string];
            nativeRouteClusterOwned(fonts: readonly [string, string], cluster: string): boolean | undefined;
          } }).oracle;
          const faces = await oracle.loadNativeRouteSentinels(document.fonts);
          if (!faces) return null;
          const list = `"${absent}", "Arial"`;
          const [line] = oracle.measureNativeRouteNormalLines([{ familyList: list, weight: 400, style: 'normal' }],
            document, () => true);
          if (line?.primaryFamily !== 'Arial') return null;
          // Stacked marks whose own ink could hide a sentinel bar.
          const STACKED = '\u0174' + '\u0301'.repeat(100) + '\u0323'.repeat(100) + '\ufe2e';
          const fonts = oracle.nativeRouteOwnershipFonts(line.primaryFamilyText, 400, 'normal');
          return Object.fromEntries(['A', 'e\u0301', ' ', '\u6f22', '\u2764\ufe0f', '1\ufe0f\u20e3', 'a\u1ab0', '\u041c\u1ab0', '\u0174\ufe2e', '\u2764\ufe0f\u0301', '\u{1f600}', STACKED]
            .map(cluster => [cluster, oracle.nativeRouteClusterOwned(fonts, cluster)]));
        }, ABSENT);
        if (!result) test.skip();
        const STACKED = '\u0174' + '\u0301'.repeat(100) + '\u0323'.repeat(100) + '\ufe2e';
        expect(result!['A']).toBe(true);
        expect(result![' ']).toBe(true);
        for (const cluster of ['\u6f22', '\u2764\ufe0f', '1\ufe0f\u20e3', 'a\u1ab0', '\u041c\u1ab0', '\u0174\ufe2e', '\u2764\ufe0f\u0301', '\u{1f600}', STACKED]) {
          expect(result![cluster], cluster).not.toBe(true);
        }
      } finally {
        await other.close();
      }
    }, 60_000);
  }
});
