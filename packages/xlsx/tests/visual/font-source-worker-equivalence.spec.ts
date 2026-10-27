import { readFileSync } from 'node:fs';
import { test, expect, type Page } from '@playwright/test';
import { PNG } from 'pngjs';
import pixelmatch from 'pixelmatch';
// @ts-ignore - shared plain-JS test fixture builder.
import { storedZip } from '../../../../tests/fixtures/chart-ex-packages.mjs';

// LoadOptions.fonts in XLSX main and worker mode (design §3.4, §4.3, §6,
// §12.1). Both realms build the same font-source demand from the same parsed
// workbook and run the same session, so they load only the unicode-range
// subsets the sheet paints and paint the same pixels. In worker mode the
// viewer realm also registers the faces (its geometry and hit testing measure
// text, design §4.3), so every request is made once per realm. The fonts are
// the committed synthetic fixtures from scripts/generate-fixture-font.py.

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

const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
function fixtureXlsx(family: string): Uint8Array {
  return storedZip([
    ['[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
      + '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
      + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
      + '</Types>'],
    ['_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`],
    ['xl/workbook.xml', `<workbook xmlns="${MAIN}" xmlns:r="${REL}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`],
    ['xl/_rels/workbook.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + `<Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/>`
      + `<Relationship Id="rId2" Type="${REL}/styles" Target="styles.xml"/></Relationships>`],
    ['xl/styles.xml', `<styleSheet xmlns="${MAIN}"><fonts count="1"><font><sz val="20"/><name val="${family}"/></font></fonts>`
      + '<fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders>'
      + '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs></styleSheet>'],
    ['xl/worksheets/sheet1.xml', `<worksheet xmlns="${MAIN}"><cols><col min="1" max="1" width="40" customWidth="1"/></cols><sheetData>`
      + '<row r="1" ht="30" customHeight="1"><c r="A1" t="inlineStr"><is><t>Fixture text ABC 123</t></is></c></row>'
      + '<row r="2" ht="30" customHeight="1"><c r="A2" t="inlineStr"><is><t>kanji 一二三</t></is></c></row>'
      + '</sheetData></worksheet>'],
  ]) as Uint8Array;
}

async function renderMode(page: Page, mode: 'main' | 'worker' | 'none', substitute: boolean) {
  const requests: Record<string, number> = {};
  const count = (key: string) => { requests[key] = (requests[key] ?? 0) + 1; };
  const bytes = fixtureXlsx(substitute ? 'Absent Office Face' : 'Ooxml Fixture');
  await page.route('**/generated/font-source.xlsx', (route) => route.fulfill({
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
  await page.goto(`/tests/visual/font-source-fixture.html?mode=${mode}${substitute ? '&substitute=1' : ''}`);
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

for (const substitute of [false, true]) {
  test(`fonts: main and worker mode request the same subsets and paint alike${substitute ? ' (substitute)' : ''}`, async ({ browser }) => {
    const results = [];
    for (const mode of ['main', 'worker', 'none'] as const) {
      // A fresh context per mode: no shared HTTP cache, FontFaceSet or realm.
      const context = await browser.newContext();
      try {
        results.push(await renderMode(await context.newPage(), mode, substitute));
      } finally {
        await context.close();
      }
    }
    const [main, worker, none] = results;
    console.log(`  font requests: main ${JSON.stringify(main.requests)}, worker ${JSON.stringify(worker.requests)}`);
    // One stylesheet, one request per demanded subset; the kana and Hangul
    // subsets are never requested (none is painted). Worker mode makes the
    // same requests in each of its two painting realms.
    expect(main.requests).toEqual({ stylesheet: 1, latin: 1, kanji: 1 });
    expect(worker.requests).toEqual({ stylesheet: 2, latin: 2, kanji: 2 });
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
}
