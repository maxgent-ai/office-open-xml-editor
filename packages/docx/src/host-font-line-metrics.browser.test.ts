import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { chromium, type Browser, type Page } from '@playwright/test';
import { createServer, type ViteDevServer } from 'vite';
import { existsSync } from 'node:fs';
import { storeZip } from './conformance/generate.js';

type HostApi = typeof import('./host-font-line-metrics') & typeof import('./layout-runtime')
  & typeof import('./line-layout') & typeof import('./index') & typeof import('@silurus/ooxml-core');

describe.skipIf(!existsSync(chromium.executablePath()))('DOCX prepared host geometry', () => {
  let server: ViteDevServer;
  let browser: Browser;
  let page: Page;

  beforeAll(async () => {
    server = await createServer({
      configFile: new URL('../vite.config.ts', import.meta.url).pathname,
      server: { host: '127.0.0.1', port: 6684, strictPort: true },
      logLevel: 'error',
    });
    await server.listen();
    browser = await chromium.launch();
    page = await browser.newPage();
    await page.addInitScript(() => { (window as unknown as { __OOXML_MODEL_SOURCES__: boolean }).__OOXML_MODEL_SOURCES__ = true; });
    await page.route('**/host-metrics.html', route => route.fulfill({
      contentType: 'text/html', body: '<!doctype html><body></body>',
    }));
    await page.goto('http://127.0.0.1:6684/host-metrics.html');
    await page.addScriptTag({ type: 'module', content: `
      import * as host from '/src/host-font-line-metrics.ts';
      import * as runtime from '/src/layout-runtime.ts';
      import * as lines from '/src/line-layout.ts';
      import { DocxDocument } from '/src/index.ts';
      import { unloadOfficeFontFallbacks } from ${JSON.stringify('/@fs' + new URL('../../core/src/index.ts', import.meta.url).pathname)};
      window.hostMetricTest = { ...host, ...runtime, ...lines, DocxDocument, unloadOfficeFontFallbacks };
    ` });
    await page.waitForFunction(() => 'hostMetricTest' in window);
  }, 60_000);
  afterAll(async () => { await browser?.close(); await server?.close(); });

  it('admits host geometry only for glyphs drawn by that source, including script slots', async (test) => {
    const result = await page.evaluate(async () => {
      const { loadDocxOfficeFontFallbacks, createLayoutServices, buildSegments, layoutLines,
        unloadOfficeFontFallbacks } = (window as unknown as { hostMetricTest: HostApi }).hostMetricTest;
      const loaded = await loadDocxOfficeFontFallbacks([{ family: 'Times New Roman' }]);
      const route = loaded.routes['times new roman'];
      if (!route?.metric.lineHeightRatio) { unloadOfficeFontFallbacks(loaded.faces); return null; }
      const ctx = document.createElement('canvas').getContext('2d') as CanvasRenderingContext2D;
      const model = {
        section: { pageWidth: 612, pageHeight: 792, marginTop: 72, marginRight: 72,
          marginBottom: 72, marginLeft: 72, headerDistance: 36, footerDistance: 36 },
        body: [], headers: {}, footers: {},
      } as unknown as import('./types').DocxDocumentModel;
      // Independent comparison: same paint route, without the new vertical
      // projection. Missing glyphs must retain this native measurement path.
      const nativeRoute = { ...route, source: 'substitute' as const, metric: { family: route.family } };
      const services = [route, nativeRoute].map(value => createLayoutServices(model, {
        measureContext: ctx, officeRoutes: [value],
      }));
      const values = ['ABC', '漢字', '😀', 'अक्षर', '\u20dd', 'ABC漢字😀अक्षर'].map(text => {
        const run = { type: 'text' as const, text, fontFamily: 'Times New Roman',
          fontFamilyHighAnsi: 'Times New Roman', fontFamilyEastAsia: 'Times New Roman',
          fontFamilyCs: 'Times New Roman', fontSize: 20, bold: false, italic: false,
          underline: false, strikethrough: false, color: null, isLink: false,
          background: null, vertAlign: null, hyperlink: null };
        return services.map(layoutServices => {
          const segments = buildSegments([run], { pageIndex: 0, totalPages: 1, layoutServices });
          const lines = layoutLines(ctx, segments, 400, 0, 1);
          return {
            text, owned: segments.filter(segment => 'text' in segment).map(segment => ({
              text: segment.text, resource: segment.resolvedResourceVerticalMetric === true,
            })),
            sides: lines.map(line => [line.ascent, line.descent, line.height]),
          };
        });
      });
      unloadOfficeFontFallbacks(loaded.faces);
      return { values };
    });
    if (!result) test.skip();
    expect(result!.values[0][0].owned.every(value => value.resource)).toBe(true);
    for (const [projected, native] of result!.values.slice(1, -1)) {
      expect(projected.owned.every(value => !value.resource)).toBe(true);
      expect(projected.sides).toEqual(native.sides);
    }
    const mixed = result!.values.at(-1)![0].owned;
    expect(mixed.some(value => value.text.includes('ABC') && value.resource)).toBe(true);
    expect(mixed.filter(value => !value.text.includes('ABC')).every(value => !value.resource)).toBe(true);
  });

  it('uses the host line geometry in both main and worker pagination', async (test) => {
    const encoder = new TextEncoder();
    const parts = new Map<string, Uint8Array>();
    parts.set('[Content_Types].xml', encoder.encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'));
    parts.set('_rels/.rels', encoder.encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'));
    const paragraph = '<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="22"/></w:rPr><w:t>Host resource pagination</w:t></w:r></w:p>';
    parts.set('word/document.xml', encoder.encode(`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${paragraph.repeat(140)}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`));
    const bytes = Array.from(storeZip(parts));
    const result = await page.evaluate(async bytes => {
      const { loadDocxOfficeFontFallbacks, DocxDocument, unloadOfficeFontFallbacks }
        = (window as unknown as { hostMetricTest: HostApi }).hostMetricTest;
      const loaded = await loadDocxOfficeFontFallbacks([{ family: 'Times New Roman' }]);
      const metric = loaded.routes['times new roman']?.metric;
      if (!metric?.lineHeightRatio) { unloadOfficeFontFallbacks(loaded.faces); return null; }
      const values = [];
      for (const mode of ['main', 'worker'] as const) {
        const doc = await DocxDocument.load(new Uint8Array(bytes).buffer, { mode });
        try {
          const runs = [];
          for (let index = 0; index < doc.pageCount; index++) runs.push(await doc.collectPageRuns(index));
          values.push({ pages: doc.pageCount, runs });
        } finally { doc.destroy(); }
      }
      unloadOfficeFontFallbacks(loaded.faces);
      return { values, pitch: metric.lineHeightRatio! * 11 * 4 / 3 };
    }, bytes);
    if (!result) test.skip();
    expect(result!.values[1]).toEqual(result!.values[0]);
    expect(result!.values[0].runs[0][0].h).toBeCloseTo(result!.pitch, 8);
  }, 60_000);

});
