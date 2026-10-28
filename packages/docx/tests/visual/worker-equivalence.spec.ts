import { test, expect } from '@playwright/test';
import { PNG } from 'pngjs';
import pixelmatch from 'pixelmatch';
import { storeZip } from '../../src/conformance/generate.js';
// @ts-ignore - shared plain-JS test fixture builder.
import { chartExDocxBytes } from '../../../../tests/fixtures/chart-ex-packages.mjs';

// Worker mode must produce (near-)identical pixels to main mode: same
// renderer, same pagination, same fonts, different thread. The only expected
// drift is sub-pixel text rasterization between the main-thread <canvas> and
// the worker OffscreenCanvas. Per-page tolerances grant slack only to a page
// that needs it, so the (near-)bit-identical pages keep enough sensitivity
// that even a single dropped text element fails the diff.
const PAGES = [0, 1];
const MAX_DIFF_PCT = [0.2, 0.2];

for (const pageIndex of PAGES) {
  test(`worker mode matches main mode › demo/sample-1 page ${pageIndex + 1}`, async ({ page }) => {
    await page.goto(`/tests/visual/worker-fixture.html?docx=demo/sample-1&page=${pageIndex}`);
    // Two full loads (main + worker) plus worker spin-up per test — twice the
    // single-render budget visual.spec.ts uses.
    await page.waitForFunction(
      () => document.body.dataset.status === 'ready' || document.body.dataset.status === 'error',
      { timeout: 60_000 },
    );
    const status = await page.evaluate(() => document.body.dataset.status);
    if (status === 'error') {
      throw new Error(await page.evaluate(() => document.body.dataset.errorMessage ?? ''));
    }

    const [mainUrl, workerUrl] = await page.evaluate(() => [
      (document.getElementById('main-canvas') as HTMLCanvasElement).toDataURL('image/png'),
      (document.getElementById('worker-canvas') as HTMLCanvasElement).toDataURL('image/png'),
    ]);
    const a = PNG.sync.read(Buffer.from(mainUrl.split(',')[1], 'base64'));
    const b = PNG.sync.read(Buffer.from(workerUrl.split(',')[1], 'base64'));
    // A zero-size canvas means a silently failed render; fail with a readable
    // assertion instead of the NaN the diff percentage would produce.
    expect(a.width).toBeGreaterThan(0);
    expect(a.height).toBeGreaterThan(0);
    expect(b.width).toBe(a.width);
    expect(b.height).toBe(a.height);

    const diff = pixelmatch(a.data, b.data, undefined, a.width, a.height, { threshold: 0.1 });
    const pct = (diff / (a.width * a.height)) * 100;
    console.log(`  page ${pageIndex + 1}: worker-vs-main diff ${pct.toFixed(3)}%`);
    expect(pct).toBeLessThanOrEqual(MAX_DIFF_PCT[pageIndex]);
  });
}

test('worker mode renders non-empty equation pixels through the same math renderer', async ({ page }) => {
  await page.goto('/tests/visual/math-worker-equivalence.html');
  await page.waitForFunction(
    () => document.body.dataset.status === 'ready' || document.body.dataset.status === 'error',
    { timeout: 60_000 },
  );
  const status = await page.evaluate(() => document.body.dataset.status);
  if (status === 'error') {
    throw new Error(await page.evaluate(() => document.body.dataset.errorMessage ?? ''));
  }

  const [mainUrl, workerUrl] = await page.evaluate(() => [
    (document.getElementById('main-canvas') as HTMLCanvasElement).toDataURL('image/png'),
    (document.getElementById('worker-canvas') as HTMLCanvasElement).toDataURL('image/png'),
  ]);
  const main = PNG.sync.read(Buffer.from(mainUrl.split(',')[1], 'base64'));
  const worker = PNG.sync.read(Buffer.from(workerUrl.split(',')[1], 'base64'));
  expect(worker.width).toBe(main.width);
  expect(worker.height).toBe(main.height);
  let inkPixels = 0;
  for (let offset = 0; offset < main.data.length; offset += 4) {
    if (main.data[offset] < 240
      || main.data[offset + 1] < 240
      || main.data[offset + 2] < 240) {
      inkPixels++;
    }
  }
  expect(inkPixels).toBeGreaterThan(500);

  const diff = pixelmatch(
    main.data,
    worker.data,
    undefined,
    main.width,
    main.height,
    { threshold: 0.1 },
  );
  const pct = (diff / (main.width * main.height)) * 100;
  console.log(`  math renderer: worker-vs-main diff ${pct.toFixed(3)}%, ink ${inkPixels}`);
  expect(pct).toBeLessThanOrEqual(0.01);
});

// The optional ChartEx renderer crosses the worker boundary as a descriptor.
// Public demo files carry no ChartEx part, so serve a self-authored waterfall
// and require identical, coloured (not grayscale-placeholder) output.
test('worker mode matches main mode › ChartEx waterfall', async ({ page }) => {
  const bytes = chartExDocxBytes();
  await page.route('**/generated/chart-ex.docx', (route) => route.fulfill({
    status: 200, contentType: 'application/octet-stream', body: Buffer.from(bytes),
  }));
  await page.goto('/tests/visual/worker-fixture.html?docx=generated/chart-ex&page=0');
  await page.waitForFunction(
    () => document.body.dataset.status === 'ready' || document.body.dataset.status === 'error',
    { timeout: 60_000 },
  );
  const status = await page.evaluate(() => document.body.dataset.status);
  if (status === 'error') {
    throw new Error(await page.evaluate(() => document.body.dataset.errorMessage ?? ''));
  }
  const [mainUrl, workerUrl] = await page.evaluate(() => [
    (document.getElementById('main-canvas') as HTMLCanvasElement).toDataURL('image/png'),
    (document.getElementById('worker-canvas') as HTMLCanvasElement).toDataURL('image/png'),
  ]);
  const main = PNG.sync.read(Buffer.from(mainUrl.split(',')[1], 'base64'));
  const worker = PNG.sync.read(Buffer.from(workerUrl.split(',')[1], 'base64'));
  expect(worker.width).toBe(main.width);
  expect(worker.height).toBe(main.height);
  const coloured = (png: PNG): number => {
    let count = 0;
    for (let offset = 0; offset < png.data.length; offset += 4) {
      const r = png.data[offset], g = png.data[offset + 1], b = png.data[offset + 2];
      if (Math.max(r, g, b) - Math.min(r, g, b) > 60) count++;
    }
    return count;
  };
  expect(coloured(main)).toBeGreaterThan(500);
  expect(coloured(worker)).toBeGreaterThan(500);
  const diff = pixelmatch(main.data, worker.data, undefined, main.width, main.height, { threshold: 0.1 });
  expect((diff / (main.width * main.height)) * 100).toBeLessThanOrEqual(0.1);
});

// wp:extent is CT_PositiveSize2D: either axis may be zero. Exercise actual
// parser acquisition, page-start square-wrap prescan and paint in both realms;
// pixel equality alone could miss both renderers silently omitting the lines.
test('worker mode paints square-wrapped zero-extent horizontal and vertical lines', async ({ page }) => {
  const line = (id: number, cx: number, cy: number, color: string) => `
    <w:r><w:drawing>
      <wp:anchor simplePos="0" relativeHeight="${id}" behindDoc="0" locked="0"
          layoutInCell="1" allowOverlap="1" distT="0" distB="0" distL="0" distR="0">
        <wp:simplePos x="0" y="0"/>
        <wp:positionH relativeFrom="page"><wp:posOffset>914400</wp:posOffset></wp:positionH>
        <wp:positionV relativeFrom="page"><wp:posOffset>1828800</wp:posOffset></wp:positionV>
        <wp:extent cx="${cx}" cy="${cy}"/>
        <wp:effectExtent l="0" t="0" r="0" b="0"/>
        <wp:wrapSquare wrapText="bothSides"/>
        <wp:docPr id="${id}" name="Line ${id}"/>
        <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
          <wps:wsp><wps:cNvSpPr/><wps:spPr>
            <a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>
            <a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:noFill/>
            <a:ln w="25400"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></a:ln>
          </wps:spPr><wps:bodyPr/></wps:wsp>
        </a:graphicData></a:graphic>
      </wp:anchor>
    </w:drawing></w:r>`;
  const xml = new TextEncoder();
  const bytes = storeZip(new Map([
    ['[Content_Types].xml', xml.encode(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
    </Types>`)],
    ['_rels/.rels', xml.encode(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
    </Relationships>`)],
    ['word/document.xml', xml.encode(`<w:document
      xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
      xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
      xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
      <w:body><w:p><w:r><w:t>Text before the lines</w:t></w:r></w:p>
        <w:p>${line(1, 1828800, 0, 'FF0000')}${line(2, 0, 914400, '0000FF')}</w:p>
        <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr>
      </w:body></w:document>`)],
  ]));
  await page.route('**/generated/zero-extent-lines.docx', (route) => route.fulfill({
    contentType: 'application/octet-stream', body: Buffer.from(bytes),
  }));
  await page.goto('/tests/visual/worker-fixture.html?docx=generated/zero-extent-lines&page=0');
  await page.waitForFunction(() => ['ready', 'error'].includes(document.body.dataset.status ?? ''));
  expect(await page.evaluate(() => document.body.dataset.errorMessage)).toBeUndefined();
  const urls = await page.locator('canvas').evaluateAll((canvases) =>
    canvases.map((canvas) => (canvas as HTMLCanvasElement).toDataURL('image/png')));
  const [main, worker] = urls.map((url) => PNG.sync.read(Buffer.from(url.split(',')[1], 'base64')));
  expect(worker.width).toBe(main.width);
  expect(worker.height).toBe(main.height);
  for (const png of [main, worker]) {
    let red = 0, blue = 0;
    for (let i = 0; i < png.data.length; i += 4) {
      if (png.data[i] > 200 && png.data[i + 1] < 50 && png.data[i + 2] < 50) red++;
      if (png.data[i] < 50 && png.data[i + 1] < 50 && png.data[i + 2] > 200) blue++;
    }
    expect(red).toBeGreaterThan(500);
    expect(blue).toBeGreaterThan(250);
  }
  const diff = pixelmatch(main.data, worker.data, undefined, main.width, main.height, { threshold: 0.1 });
  expect(diff / (main.width * main.height)).toBeLessThanOrEqual(0.001);
});
