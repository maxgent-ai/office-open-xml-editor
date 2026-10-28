import { readFileSync } from 'node:fs';
import { test, expect, type Page } from '@playwright/test';
import { PNG } from 'pngjs';
import pixelmatch from 'pixelmatch';
import { deck, para, textBox } from '../../src/test-support/synthetic-deck.js';

// LoadOptions.fonts in PPTX main and worker mode (design §3.4, §6, §12.1).
// Both realms build the same font-source demand from the same preflight facts
// and run the same session, so they must request the same stylesheet and font
// files, load only the unicode-range subsets the slides paint, and paint the
// same pixels, in ordinary and progressive loads. The fonts are the committed
// synthetic fixtures from scripts/generate-fixture-font.py (no third-party
// copyright).

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

const fixtureRun = (family: string, text: string) => `<a:r><a:rPr lang="en-US" sz="4000"><a:latin typeface="${family}"/>`
  + `<a:ea typeface="${family}"/><a:cs typeface="${family}"/></a:rPr><a:t>${text}</a:t></a:r>`;
const fixtureDeck = (family: string) =>
  deck(textBox(para(fixtureRun(family, 'Fixture text ABC 123')) + para(fixtureRun(family, 'kanji 一二三'))));

interface Variant { readonly progressive: boolean; readonly substitute: boolean | 'bullet' }

// Latin body text; an Arabic character bullet (§21.1.2.4.3 buChar) in a face
// the host lacks, served only by an Arabic-scoped substitute.
const BULLET_PPTX = deck(textBox(para(fixtureRun('Body Face', 'Latin body'),
  '<a:pPr marL="342900" indent="-342900"><a:buFont typeface="Absent Bullet Face"/><a:buChar char="ا"/></a:pPr>')));

async function renderMode(page: Page, mode: 'main' | 'worker' | 'none', { progressive, substitute }: Variant) {
  const FIXTURE_PPTX = substitute === 'bullet' ? BULLET_PPTX : fixtureDeck(substitute ? 'Absent Office Face' : 'Ooxml Fixture');
  const requests: Record<string, number> = {};
  const count = (key: string) => { requests[key] = (requests[key] ?? 0) + 1; };
  await page.route('**/generated/font-source.pptx', (route) => route.fulfill({
    status: 200, contentType: 'application/octet-stream', body: Buffer.from(FIXTURE_PPTX),
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
  await page.goto(`/tests/visual/font-source-fixture.html?mode=${mode}${progressive ? '&progressive=1' : ''}${substitute === 'bullet' ? '&substitute=bullet' : substitute ? '&substitute=1' : ''}`);
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

// `substitute`: the deck's face is served only as a declared substitute, so
// the renderer's font stack must insert the loaded substitution in both realms.
const VARIANTS: readonly Variant[] = [
  { progressive: false, substitute: false },
  { progressive: true, substitute: false },
  { progressive: false, substitute: true },
];

for (const variant of VARIANTS) {
  const label = [variant.progressive && 'progressive', variant.substitute && 'substitute'].filter(Boolean).join(', ');
  test(`fonts: main and worker mode request the same subsets and paint alike${label ? ` (${label})` : ''}`, async ({ browser }) => {
    const results = [];
    for (const mode of ['main', 'worker', 'none'] as const) {
      // A fresh context per mode: no shared HTTP cache, FontFaceSet or realm.
      const context = await browser.newContext();
      try {
        results.push(await renderMode(await context.newPage(), mode, variant));
      } finally {
        await context.close();
      }
    }
    const [main, worker, none] = results;
    console.log(`  font requests: main ${JSON.stringify(main.requests)}, worker ${JSON.stringify(worker.requests)}`);
    // One stylesheet, one request per demanded subset; the kana and Hangul
    // subsets are never requested (none is painted), identically in both realms.
    expect(main.requests).toEqual(worker.requests);
    expect(main.requests).toEqual({ stylesheet: 1, latin: 1, kanji: 1 });
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

test('fonts: a marker\'s script proves its scoped substitute in main and worker mode', async ({ browser }) => {
  const results = [];
  for (const mode of ['main', 'worker'] as const) {
    const context = await browser.newContext();
    try {
      results.push(await renderMode(await context.newPage(), mode, { progressive: false, substitute: 'bullet' }));
    } finally {
      await context.close();
    }
  }
  const [main, worker] = results;
  console.log(`  bullet font requests: main ${JSON.stringify(main.requests)}, worker ${JSON.stringify(worker.requests)}`);
  expect(main.requests).toEqual(worker.requests);
  // The Arabic marker text proves the bullet face's Arabic-scoped substitute,
  // so the served family's faces load (bullets make scalar demand 'all').
  const files = Object.entries(main.requests).filter(([key]) => key !== 'stylesheet');
  expect(files.length).toBeGreaterThan(0);
});
