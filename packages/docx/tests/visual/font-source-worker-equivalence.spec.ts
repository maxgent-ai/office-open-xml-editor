import { readFileSync } from 'node:fs';
import { test, expect, type Page } from '@playwright/test';
import { PNG } from 'pngjs';
import pixelmatch from 'pixelmatch';
import { storeZip } from '../../src/conformance/generate.js';

// LoadOptions.fonts in DOCX main and worker mode (design §3.4, §6, §12.1).
// Both realms run the same font-source session from the same model, so they
// must request the same stylesheet and font files, load only the
// unicode-range subsets the document paints, and paint the same pixels. The
// fonts are the committed synthetic fixtures from
// scripts/generate-fixture-font.py (no third-party copyright).

const FONT_DIR = new URL('../../../../tests/fixtures/fonts/', import.meta.url);
const LATIN = readFileSync(new URL('ooxml-fixture.ttf', FONT_DIR));
const CJK = readFileSync(new URL('ooxml-fixture-jp.ttf', FONT_DIR));
const FILES: Record<string, Buffer> = { latin: LATIN, kana: CJK, kanji: CJK, hangul: CJK };
const CSS = [
  ['latin', 'U+0-FF'],
  ['kana', 'U+3000-30FF'],
  ['kanji', 'U+4E00-4FFF'],
  ['hangul', 'U+AC00-ACFF'],
].map(([name, range]) => `@font-face { font-family: 'Ooxml Fixture'; font-weight: 400;`
  + ` src: url(fonts/${name}.ttf) format('truetype'); unicode-range: ${range}; }`).join('\n');

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const encoder = new TextEncoder();
function fixtureDocx(): Uint8Array {
  const run = (text: string) => '<w:r><w:rPr><w:rFonts w:ascii="Ooxml Fixture" w:hAnsi="Ooxml Fixture"'
    + ' w:eastAsia="Ooxml Fixture" w:cs="Ooxml Fixture"/><w:sz w:val="48"/></w:rPr>'
    + `<w:t xml:space="preserve">${text}</w:t></w:r>`;
  return storeZip(new Map([
    ['[Content_Types].xml', encoder.encode('<?xml version="1.0" encoding="UTF-8"?>'
      + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      + '</Types>')],
    ['_rels/.rels', encoder.encode('<?xml version="1.0" encoding="UTF-8"?>'
      + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
      + '</Relationships>')],
    ['word/document.xml', encoder.encode('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + `<w:document xmlns:w="${W}"><w:body>`
      + `<w:p>${run('Fixture text ABC 123')}</w:p><w:p>${run('kanji 一二三')}</w:p>`
      + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
      + '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
      + '</w:sectPr></w:body></w:document>')],
  ]));
}

async function renderMode(page: Page, mode: 'main' | 'worker' | 'none') {
  const requests: Record<string, number> = {};
  const count = (key: string) => { requests[key] = (requests[key] ?? 0) + 1; };
  const bytes = fixtureDocx();
  await page.route('**/generated/font-source.docx', (route) => route.fulfill({
    status: 200, contentType: 'application/octet-stream', body: Buffer.from(bytes),
  }));
  await page.route('**/generated/fixture-fonts.css', (route) => {
    count('stylesheet');
    return route.fulfill({ status: 200, contentType: 'text/css', body: CSS });
  });
  await page.route('**/generated/fonts/*.ttf', (route) => {
    const name = /fonts\/(\w+)\.ttf$/.exec(route.request().url())?.[1] ?? '?';
    count(name);
    return route.fulfill({ status: 200, contentType: 'font/ttf', body: FILES[name] });
  });
  await page.goto(`/tests/visual/font-source-fixture.html?mode=${mode}`);
  await page.waitForFunction(
    () => document.body.dataset.status === 'ready' || document.body.dataset.status === 'error',
    { timeout: 60_000 },
  );
  if (await page.evaluate(() => document.body.dataset.status) === 'error') {
    throw new Error(await page.evaluate(() => document.body.dataset.errorMessage ?? ''));
  }
  const url = await page.evaluate(() => (document.getElementById('canvas') as HTMLCanvasElement).toDataURL('image/png'));
  return { requests, png: PNG.sync.read(Buffer.from(url.split(',')[1], 'base64')) };
}

const diffPct = (a: PNG, b: PNG) =>
  (pixelmatch(a.data, b.data, undefined, a.width, a.height, { threshold: 0.1 }) / (a.width * a.height)) * 100;

test('fonts: main and worker mode request the same subsets and paint alike', async ({ browser }) => {
  const results = [];
  for (const mode of ['main', 'worker', 'none'] as const) {
    // A fresh context per mode: no shared HTTP cache, FontFaceSet or realm.
    const context = await browser.newContext();
    try {
      results.push(await renderMode(await context.newPage(), mode));
    } finally {
      await context.close();
    }
  }
  const [main, worker, none] = results;
  console.log(`  font requests: main ${JSON.stringify(main.requests)}, worker ${JSON.stringify(worker.requests)}`);
  // One stylesheet, one request per demanded subset; the Hangul subset is
  // never requested (no Hangul is painted), identically in both realms.
  expect(main.requests).toEqual(worker.requests);
  expect(main.requests).toMatchObject({ stylesheet: 1, latin: 1, kanji: 1 });
  expect(main.requests.hangul).toBeUndefined();
  // Without a source nothing is requested.
  expect(none.requests).toEqual({});

  expect(worker.png.width).toBe(main.png.width);
  expect(worker.png.height).toBe(main.png.height);
  const mainVsWorker = diffPct(main.png, worker.png);
  const mainVsNone = diffPct(main.png, none.png);
  console.log(`  worker-vs-main diff ${mainVsWorker.toFixed(3)}%, fixture-vs-fallback ${mainVsNone.toFixed(3)}%`);
  expect(mainVsWorker).toBeLessThanOrEqual(0.2);
  // The fixture face actually painted (it differs visibly from host fallback).
  expect(mainVsNone).toBeGreaterThan(0.5);
});
